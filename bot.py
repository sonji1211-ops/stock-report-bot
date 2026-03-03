import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import time

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def fetch_stock_data(row, start_d, end_d, semaphore):
    """야후 파이낸스 소스를 사용하여 데이터를 안전하게 가져옵니다."""
    async with semaphore:
        try:
            suffix = ".KS" if row['Market'] == 'KOSPI' else ".KQ"
            ticker = f"{row['Code']}{suffix}"
            df = fdr.DataReader(ticker, start_d, end_d)
            if df is None or len(df) < 2:
                return None
            
            last_c = float(df.iloc[-1]['Close'])
            prev_c = float(df.iloc[-2]['Close'])
            ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
            
            return {
                'Code': row['Code'], 'Name': row['Name'], 
                'Open': df.iloc[-1]['Open'], 'Close': last_c,
                'Low': df['Low'].min(), 'High': df['High'].max(), 
                'Ratio': ratio, 'Volume': df.iloc[-1]['Volume'],
                'Market': row['Market']
            }
        except:
            return None

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty:
            return

        sem = asyncio.Semaphore(15)
        
        if day_of_week == 6: # 일요일: 주간 500개
            report_type, analysis_info = "주간평균", "시총 상위 500"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
        else: # 평일: 전 종목 전수조사
            report_type, analysis_info = "일일", "전 종목 전수조사"
            if day_of_week == 5:
                report_type = "일일(금요마감)"
            end_d = now.strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=5)).strftime('%Y-%m-%d')
            df_target = df_base.copy()

        tasks = [fetch_stock_data(row, start_d, end_d, sem) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        df_final = pd.DataFrame([r for r in results if r is not None])
        
        if df_final.empty:
            return

        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_data(m_name, is_up):
            temp = df_final[df_final['Market'].str.contains(m_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {
            '코스피_상승': get_data('KOSPI', True), 
            '코스닥_상승': get_data('KOSDAQ', True),
            '코스피_하락': get_data('KOSPI', False), 
            '코스닥_하락': get_data('KOSDAQ', False)
        }

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        header_fill = PatternFill("solid", fgColor="444444")
        header_font = Font(color="FFFFFF", bold=True)
        fills = [
            PatternFill("solid", fgColor="FF0000"), # 빨강
            PatternFill("solid", fgColor="FFBB00"), # 주황
            PatternFill("solid", fgColor="FFFF00")  # 노랑
        ]
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                for row in range(2, ws.max_row + 1):
                    ratio = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    
                    if ratio >= 28:
                        name_cell.fill, name_cell.font = fills[0], Font(color="FFFFFF", bold=True)
                    elif ratio >= 20:
                        name_cell.fill = fills[1]
                    elif ratio >= 10:
                        name_cell.fill = fills[2]
                    
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = border
                        if col in [3,4,5,6,8]:
                            ws.cell(row, col).number_format = '#,##0'
                        if col == 7:
                            ws.cell(row, col).number_format = '0.00'

                ws.column_dimensions['B'].width = 20
                for char in "ACDEFGH":
                    ws.column_dimensions[char].width = 12

        async with bot:
            date_str = f"{start_d}~{end_d}" if day_of_week == 6 else now.strftime('%Y-%m-%d')
            msg = (f"📦 *[{report_type}] 주식 리포트 도착*\n\n"
                   f"📅 *날짜:* {date_str}\n"
                   f"🔍 *대상:* {analysis_info}\n"
                   f"───\n"
                   f"📈 *상승(5%↑):* {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])}개\n"
                   f"📉 *하락(5%↓):* {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}개\n\n"
                   f"💡 *강조 안내*\n"
                   f"🔴 28%↑ | 🟠 20%↑ | 🟡 10%↑")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg, parse_mode="Markdown")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
