import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    # 1. 한국 시간 설정 및 주말 보정
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 
    
    # 보고서 성격 정의
    if day_of_week == 6: # 일요일 실행 (월~금 누적 데이터)
        report_type = "주간누적(월~금)"
        end_date = (now - timedelta(days=2)).strftime('%Y-%m-%d')
        start_date = (now - timedelta(days=6)).strftime('%Y-%m-%d')
    elif day_of_week == 5: # 토요일 실행 (금요일 마감 데이터)
        report_type = "일일(금요일마감)"
        end_date = (now - timedelta(days=1)).strftime('%Y-%m-%d')
        start_date = end_date
    else: # 평일
        report_type = "일일"
        end_date = now.strftime('%Y-%m-%d')
        start_date = end_date

    try:
        print(f"--- 분석 모드: {report_type} ---")
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        # 2. 데이터 가공 (일요일은 누적 / 그외는 당일)
        if day_of_week == 6:
            print("일요일 주간 평균 데이터를 수집 중입니다...")
            weekly_data = []
            df_target = df_base.sort_values(by='Volume', ascending=False).head(1500)
            for idx, row in df_target.iterrows():
                try:
                    d_hist = fdr.DataReader(row['Code'], start_date, end_date)
                    if len(d_hist) >= 2:
                        open_p, close_p = d_hist.iloc[0]['Open'], d_hist.iloc[-1]['Close']
                        ratio = round(((close_p - open_p) / open_p) * 100, 2)
                        weekly_data.append({
                            'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                            'Open': open_p, 'High': d_hist['High'].max(), 'Low': d_hist['Low'].min(),
                            'Close': close_p, 'Calculated_Ratio': ratio, 'Volume': d_hist['Volume'].mean()
                        })
                except: continue
            df = pd.DataFrame(weekly_data)
        else:
            cols = df_base.columns.tolist()
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), None)
            df_base['Calculated_Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
            if df_base['Calculated_Ratio'].abs().max() < 2: df_base['Calculated_Ratio'] *= 100
            df = df_base.copy()
            df['Calculated_Ratio'] = df['Calculated_Ratio'].round(2)

        # 3. 엑셀 구조 잡기
        h_map = {'Code': '종목코드', 'Name': '종목명', 'Market': '시장', 'Open': '시가', 
                 'High': '고가', 'Low': '저가', 'Close': '종가', 'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'}

        def process_data(market, is_up):
            m_df = df[df['Market'].str.contains(market, na=False)].copy()
            res = m_df[m_df['Calculated_Ratio'] >= 5] if is_up else m_df[m_df['Calculated_Ratio'] <= -5]
            res = res.sort_values(by='Calculated_Ratio', ascending=not is_up)
            return res[[c for c in h_map.keys() if c in res.columns]].rename(columns=h_map)

        sheets_data = {'코스피_상승': process_data('KOSPI', True), '코스닥_상승': process_data('KOSDAQ', True),
                       '코스피_하락': process_data('KOSPI', False), '코스닥_하락': process_data('KOSDAQ', False)}

        # 4. 엑셀 파일 생성 및 색상(28% 기준) 입히기
        file_name = f"{now.strftime('%Y-%m-%d')}_{report_type}_국내리포트.xlsx"
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row=row, column=8).value or 0))
                    name_cell = ws.cell(row=row, column=2)
                    
                    # 지수님 요청 색상 임계값 (10/20/28)
                    if val >= 28: 
                        name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: 
                        name_cell.fill = fill_orange
                    elif val >= 10: 
                        name_cell.fill = fill_yellow
                    
                    # 가운데 정렬 및 숫자 포맷
                    for c in range(1, 10):
                        cell = ws.cell(row=row, column=c)
                        cell.alignment = Alignment(horizontal='center')
                        if c == 8: cell.number_format = '0.00'
                        elif c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                for i in range(1, 10): ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 15

        # 5. 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀 종목명 색상 가이드\n(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
    except Exception as e: print(f"국장 에러: {e}")

if __name__ == "__main__": asyncio.run(send_smart_report())
