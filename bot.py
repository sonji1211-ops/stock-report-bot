import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    # 1. 날짜 및 분석 타겟 설정
    if day_of_week == 6: # 일요일: 주간평균 (시가총액 상위 500)
        report_type = "주간평균(시총상위)"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d')
        start_d, end_d = (now - timedelta(days=6)).strftime('%Y-%m-%d'), target_date_str
        sort_column = 'Marcap' # 시가총액 기준
    else: # 화~토: 일일 (거래량 상위 500)
        report_type = "일일"
        if day_of_week == 5: report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1 if day_of_week == 5 else 0)).strftime('%Y-%m-%d')
        start_d = end_d = target_date_str
        sort_column = 'Volume' # 거래량 기준

    try:
        print(f"--- {report_type} 분석 시작 (기준: {sort_column}) ---")
        
        # 2. 전체 종목 리스트 확보 및 타겟팅 (500개)
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return
        
        # 요일별로 정해진 기준(시총/거래량)에 따라 500개 추출
        df_target = df_base.sort_values(by=sort_column, ascending=False).head(500).copy()

        # 3. 고속 병렬 데이터 수집 함수
        async def fetch_stock(row):
            try:
                # 안전하게 7~10일치 데이터 확보
                h = fdr.DataReader(row['Code'], (datetime.strptime(start_d, '%Y-%m-%d')-timedelta(days=10)).strftime('%Y-%m-%d'), end_d)
                if h.empty or len(h) < 2: return None
                
                if day_of_week == 6: # [일요일] 월~금 일별 등락률의 '평균'
                    # 주간 범위 내에서만 수익률 계산
                    h['rt'] = h['Close'].pct_change() * 100
                    target_h = h.loc[start_d:end_d]
                    if target_h.empty: return None
                    ratio = round(target_h['rt'].mean(), 2)
                else: # [평일/토요일] 어제 종가 대비 오늘 종가
                    ratio = round(((h.iloc[-1]['Close'] - h.iloc[-2]['Close']) / h.iloc[-2]['Close']) * 100, 2)
                
                return {
                    'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                    'Open': h.iloc[-1]['Open'], 'High': h['High'].max(), 'Low': h['Low'].min(),
                    'Close': h.iloc[-1]['Close'], 'Calculated_Ratio': ratio, 'Volume': h.iloc[-1]['Volume']
                }
            except: return None

        # 4. 병렬 처리로 속도 극대화
        tasks = [fetch_stock(row) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        res_list = [r for r in results if r is not None]

        df_final = pd.DataFrame(res_list)
        if df_final.empty: return

        # 5. 분류 및 엑셀 스타일 적용
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Market':'시장', 'Open':'시가', 'High':'고가', 'Low':'저가', 'Close':'종가', 'Calculated_Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_sub(market, is_up):
            m_df = df_final[df_final['Market'].str.contains(market, na=False)].copy()
            cond = (m_df['Calculated_Ratio'] >= 5) if is_up else (m_df['Calculated_Ratio'] <= -5)
            # 엑셀에서도 등락률 순으로 정렬해서 보여줌
            return m_df[cond].sort_values('Calculated_Ratio', ascending=not is_up)[list(h_map.keys())].rename(columns=h_map)

        sheets_data = {'코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True),
                       '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)}

        file_name = f"{now.strftime('%Y-%m-%d')}_{report_type}_리포트.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row, 8).value or 0))
                    name_cell = ws.cell(row, 2)
                    if val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: name_cell.fill = fill_orange
                    elif val >= 10: name_cell.fill = fill_yellow
                    for c in range(1, 10):
                        ws.cell(row, c).alignment = Alignment(horizontal='center')
                        if c == 8: ws.cell(row, c).number_format = '0.00'
                for i in range(1, 10): ws.column_dimensions[chr(64+i)].width = 15

        # 6. 텔레그램 발송
        async with bot:
            base_msg = (f"📅 {target_date_str} {report_type} 리포트\n\n"
                        f"📊 분석기준: {'시가총액 상위 500' if day_of_week==6 else '거래량 상위 500'}\n"
                        f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                        f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                        f"💡 🟡10%↑, 🟠20%↑, 🔴28%↑")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=base_msg)

    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
