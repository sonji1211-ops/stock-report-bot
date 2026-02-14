import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font
import time

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    # 1. 날짜 및 리포트 타입 설정
    if day_of_week == 6: # 일요일 실행: 이번 주 주간 일별 등락률 평균
        report_type = "주간평균"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d') # 금요일
        start_d, end_d = (now - timedelta(days=6)).strftime('%Y-%m-%d'), target_date_str
        sample_count = 1000 # 주간은 상위 1,000개 집중 분석
    else: # 화~토 실행: 전일 대비 일일 리포트
        report_type = "일일"
        if day_of_week == 5: report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1 if day_of_week == 5 else 0)).strftime('%Y-%m-%d')
        start_d = end_d = target_date_str
        sample_count = 0 # 전 종목 분석

    try:
        print(f"--- {report_type} 분석 시작: {target_date_str} ---")
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        # 분석 대상 선정
        df_target = df_base.sort_values(by='Volume', ascending=False).head(sample_count) if sample_count > 0 else df_base.copy()
        
        res_list = []
        for idx, row in df_target.iterrows():
            try:
                # 등락률 계산을 위해 필요한 범위(전주 포함) 조회
                search_start = (datetime.strptime(start_d, '%Y-%m-%d') - timedelta(days=7)).strftime('%Y-%m-%d')
                h = fdr.DataReader(row['Code'], search_start, end_d)
                
                if not h.empty:
                    if day_of_week == 6: # [일요일] 월~금 일별 등락률의 '평균'
                        weekly_h = h.loc[start_d:end_d].copy()
                        if len(weekly_h) >= 2:
                            # 전체 데이터에서 일별 등락률을 먼저 구한 뒤 주간 범위만 추출
                            h['daily_rt'] = h['Close'].pct_change() * 100
                            avg_ratio = round(h.loc[start_d:end_d, 'daily_rt'].mean(), 2)
                            
                            res_list.append({
                                'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                                'Open': weekly_h.iloc[-1]['Open'], 'High': weekly_h['High'].max(),
                                'Low': weekly_h['Low'].min(), 'Close': weekly_h.iloc[-1]['Close'],
                                'Calculated_Ratio': avg_ratio, 'Volume': weekly_h.iloc[-1]['Volume']
                            })
                    else: # [화~토] 어제 종가 vs 오늘 종가
                        if len(h) >= 2:
                            o, c = h.iloc[-2]['Close'], h.iloc[-1]['Close']
                            ratio = round(((c - o) / o) * 100, 2)
                            res_list.append({
                                'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                                'Open': h.iloc[-1]['Open'], 'High': h.iloc[-1]['High'],
                                'Low': h.iloc[-1]['Low'], 'Close': c,
                                'Calculated_Ratio': ratio, 'Volume': h.iloc[-1]['Volume']
                            })
            except: continue
            if idx % 300 == 0: await asyncio.sleep(0.01)

        df_final = pd.DataFrame(res_list)
        if df_final.empty: return

        # 2. 데이터 분류 (5% 기준)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Market':'시장', 'Open':'시가', 'High':'고가', 'Low':'저가', 'Close':'종가', 'Calculated_Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_sub(market, is_up):
            m_df = df_final[df_final['Market'].str.contains(market, na=False)].copy()
            cond = (m_df['Calculated_Ratio'] >= 5) if is_up else (m_df['Calculated_Ratio'] <= -5)
            return m_df[cond].sort_values('Calculated_Ratio', ascending=not is_up)[list(h_map.keys())].rename(columns=h_map)

        sheets_data = {'코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True),
                       '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)}

        # 3. 엑셀 꾸미기 (🔴28%↑ 상한가 기준 강조)
        file_name = f"{now.strftime('%Y-%m-%d')}_{report_type}_리포트.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row, 8).value or 0)) # 등락률(%)
                    name_cell = ws.cell(row, 2) # 종목명
                    if val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: name_cell.fill = fill_orange
                    elif val >= 10: name_cell.fill = fill_yellow
                    for c in range(1, 10):
                        ws.cell(row, c).alignment = Alignment(horizontal='center')
                        if c == 8: ws.cell(row, c).number_format = '0.00'
                        elif c in [4, 5, 6, 7, 9]: ws.cell(row, c).number_format = '#,##0'
                for i in range(1, 10): ws.column_dimensions[chr(64+i)].width = 15

        # 4. 전송
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀에서 종목명 색깔을 확인하세요!\n"
                   f"(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
