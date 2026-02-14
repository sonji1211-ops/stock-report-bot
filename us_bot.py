import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_us_nasdaq100_full_report():
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 
    # 일요일 실행 시 미국은 아직 금요일 마감 데이터가 최신임
    report_type = "주간(평균)" if day_of_week == 6 else "일일"

    try:
        df_nas = fdr.StockListing('NASDAQ').head(100)
        report_list = []
        for idx, row in df_nas.iterrows():
            try:
                # 2일치 데이터를 가져와 등락률 계산
                df_p = fdr.DataReader(row['Symbol']).tail(2)
                if len(df_p) < 2: continue
                curr, prev = df_p.iloc[-1], df_p.iloc[-2]
                chg = round(((curr['Close'] - prev['Close']) / prev['Close']) * 100, 2)
                
                report_list.append({
                    '티커': row['Symbol'], '종목명': row['Name'], 
                    '시가($)': round(curr['Open'], 2), '고가($)': round(curr['High'], 2), 
                    '저가($)': round(curr['Low'], 2), '종가($)': round(curr['Close'], 2), 
                    '등락률(%)': chg
                })
            except: continue

        df_final = pd.DataFrame(report_list).sort_values(by='등락률(%)', ascending=False)
        file_name = f"{now.strftime('%Y-%m-%d')}_미국나스닥리포트.xlsx"
        
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            for row in range(2, ws.max_row + 1):
                val = abs(float(ws.cell(row=row, column=7).value or 0)) # 등락률은 7번열
                name_cell = ws.cell(row=row, column=2)
                
                # 미국장도 10/20/28 기준 색상 적용
                if val >= 28: 
                    name_cell.fill, name_cell.font = fill_red, font_white
                elif val >= 20: 
                    name_cell.fill = fill_orange
                elif val >= 10: 
                    name_cell.fill = fill_yellow
                    
                for col in range(1, 8):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    if col >= 3: cell.number_format = '0.00'
            ws.column_dimensions['B'].width = 28

        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"🇺🇸 {now.strftime('%Y-%m-%d')} 나스닥 리포트 배달완료!\n\n"
                   f"💡 엑셀 종목명 색상 가이드\n(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
    except Exception as e: print(f"미국장 에러: {e}")

if __name__ == "__main__": asyncio.run(send_us_nasdaq100_full_report())
