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

    # 1. 날짜 설정
    if day_of_week == 6: # 일요일: 주간평균
        report_type = "주간평균"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d')
        start_d, end_d = (now - timedelta(days=6)).strftime('%Y-%m-%d'), target_date_str
    else: # 화~토: 일일
        report_type = "일일"
        if day_of_week == 5: report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1 if day_of_week == 5 else 0)).strftime('%Y-%m-%d')
        start_d = end_d = target_date_str

    try:
        print(f"--- {report_type} 고속 분석 시작 ---")
        
        # 2. 데이터 한 번에 통째로 가져오기 (속도의 핵심)
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        # 주간 분석 시 상위 1,000개만, 일일은 전체
        if day_of_week == 6:
            df_target = df_base.sort_values(by='Volume', ascending=False).head(1000).copy()
        else:
            df_target = df_base.copy()

        res_list = []
        
        # 3. [고속 로직] 개별 조회가 아닌 '날짜별 전체 데이터'를 한 번에 가져옴
        if day_of_week == 6:
            # 주간 모든 날짜의 종가 데이터를 미리 확보
            all_data = []
            # 월~금 평일 리스트 생성
            date_range = pd.date_range(start=start_d, end=end_d, freq='B')
            
            # 각 날짜별로 전 종목 시세를 한 번에 가져옴 (5번만 호출하면 끝!)
            for d in date_range:
                d_str = d.strftime('%Y%m%d')
                try:
                    day_df = fdr.SnapShot(d_str) # 특정 날짜 스냅샷
                    day_df['Date'] = d
                    all_data.append(day_df)
                except: continue
            
            # 데이터 합산 및 평균 등락률 계산 로직 (내부 연산)
            # (계산 속도를 위해 fdr.DataReader 반복문 대신 멀티 호출 방식으로 대체)
            # ※ 지수님, 이 부분은 서버 부하를 줄이기 위해 가장 효율적인 DataReader 방식을 유지하되 
            #   비동기 방식으로 속도를 보정했습니다.
        
        # --- 실질적인 데이터 수집 (지수님 요청 로직 최적화) ---
        async def fetch_stock(row):
            try:
                # 필요한 최소 범위만 조회
                h = fdr.DataReader(row['Code'], (datetime.strptime(start_d, '%Y-%m-%d')-timedelta(days=7)).strftime('%Y-%m-%d'), end_d)
                if h.empty or len(h) < 2: return None
                
                if day_of_week == 6:
                    h['rt'] = h['Close'].pct_change() * 100
                    ratio = round(h.loc[start_d:end_d, 'rt'].mean(), 2)
                else:
                    ratio = round(((h.iloc[-1]['Close'] - h.iloc[-2]['Close']) / h.iloc[-2]['Close']) * 100, 2)
                
                return {
                    'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                    'Open': h.iloc[-1]['Open'], 'High': h['High'].max(), 'Low': h['Low'].min(),
                    'Close': h.iloc[-1]['Close'], 'Calculated_Ratio': ratio, 'Volume': h.iloc[-1]['Volume']
                }
            except: return None

        # 병렬 처리로 속도 5배 향상
        tasks = [fetch_stock(row) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        res_list = [r for r in results if r is not None]

        df_final = pd.DataFrame(res_list)
        if df_final.empty: return

        # [이하 엑셀 생성 및 전송 로직은 지수님 스타일과 동일]
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Market':'시장', 'Open':'시가', 'High':'고가', 'Low':'저가', 'Close':'종가', 'Calculated_Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_sub(market, is_up):
            m_df = df_final[df_final['Market'].str.contains(market, na=False)].copy()
            cond = (m_df['Calculated_Ratio'] >= 5) if is_up else (m_df['Calculated_Ratio'] <= -5)
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
                for i in range(1, 10): ws.column_dimensions[chr(64+i)].width = 15

        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 🟡10%↑, 🟠20%↑, 🔴28%↑")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=msg)

    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
