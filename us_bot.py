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

async def send_us_nasdaq100_detail_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    try:
        print(f"--- 나스닥 100 상세 데이터 강제 수집 시작 ---")
        
        # 1. 나스닥 종목 리스트 확보 (여기서 티커만 가져옵니다)
        df_nas = fdr.StockListing('NASDAQ')
        top_100_tickers = df_nas.head(100) # 상위 100개

        report_list = []

        # 2. 각 종목별로 '진짜 시세' 하나씩 가져오기
        for idx, row in top_100_tickers.iterrows():
            ticker = row['Symbol']
            name = row['Name']
            print(f"데이터 수집 중: {ticker} ({name})")
            
            try:
                # 최근 2일치 데이터를 가져와서 어제와 오늘 비교
                df = fdr.DataReader(ticker).tail(2)
                if len(df) < 2: continue
                
                prev_close = df.iloc[0]['Close'] # 전일 종가
                curr_close = df.iloc[1]['Close'] # 현재 마감가
                curr_open = df.iloc[1]['Open']   # 오늘 시작가
                curr_high = df.iloc[1]['High']   # 오늘 고가
                curr_low = df.iloc[1]['Low']     # 오늘 저가
                
                chg_ratio = ((curr_close - prev_close) / prev_close) * 100

                report_list.append({
                    '티커': ticker,
                    '종목명': name,
                    '시작가($)': curr_open,
                    '마감가($)': curr_close,
                    '고가($)': curr_high,
                    '저가($)': curr_low,
                    '등락률(%)': chg_ratio
                })
            except:
                print(f"{ticker} 수집 실패, 건너뜁니다.")
                continue

        # 3. 엑셀 파일 생성
        df_final = pd.DataFrame(report_list)
        file_name = f"{target_date_str}_나스닥100_상세리포트.xlsx"
        
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            
            for row in range(2, ws.max_row + 1):
                ratio_val = ws.cell(row=row, column=7).value # 등락률 컬럼
                name_cell = ws.cell(row=row, column=2)
                
                if ratio_val is not None:
                    if ratio_val > 0:
                        name_cell.fill = fill_red
                        name_cell.font = white_font
                    elif ratio_val < 0:
                        name_cell.fill = fill_blue
                        name_cell.font = white_font

                for col in range(1, 8):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

            ws.column_dimensions['B'].width = 30
            for i in range(3, 8):
                ws.column_dimensions[chr(64+i)].width = 15

        # 4. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {target_date_str} 나스닥 100 상세 리포트\n종목별 시가, 종가, 등락률이 모두 포함되었습니다."
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] 나스닥 100 상세 리포트 전송 완료 ---")

    except Exception as e:
        print(f"최종 에러: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_detail_report())
