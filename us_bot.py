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

async def send_us_all_stocks_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    try:
        print(f"--- 미국 전 종목 시세 데이터 수집 시작 ---")
        
        # 1. 시세 정보가 포함된 미국 주식 리스팅
        # 'NASDAQ', 'NYSE', 'AMEX'를 각각 가져옵니다.
        exchanges = ['NASDAQ', 'NYSE']
        frames = []

        for ex in exchanges:
            print(f"{ex} 데이터 불러오는 중...")
            df = fdr.StockListing(ex)
            if df is not None and not df.empty:
                df['Exchange'] = ex
                frames.append(df)
        
        all_df = pd.concat(frames, ignore_index=True)

        # 2. 컬럼 정리 (데이터 소스에 따라 컬럼명이 다를 수 있어 유연하게 처리)
        # FinanceDataReader의 미국 리스팅은 보통 Symbol, Name, Industry, ClosingPrice, ChgCode, ChngPct 등을 줍니다.
        h_map = {
            'Symbol': '티커',
            'Name': '종목명',
            'Industry': '산업',
            'Close': '종가($)',
            'Open': '시가($)',
            'High': '고가($)',
            'Low': '저가($)',
            'ChgPct': '등락률(%)',
            'Exchange': '거래소'
        }
        
        # 만약 fdr에서 주는 컬럼명이 'Close'가 아니라 'Price'라면 맞춰줍니다.
        all_df = all_df.rename(columns={'Price': 'Close', 'ChangesRatio': 'ChgPct'})
        
        final_df = all_df[[c for c in h_map.keys() if c in all_df.columns]].copy()
        final_df = final_df.rename(columns=h_map)

        # 3. 엑셀 파일 생성
        file_name = f"{target_date_str}_미국_전종목_시세.xlsx"
        
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            final_df.to_excel(writer, sheet_name='미국주식시세', index=False)
            ws = writer.sheets['미국주식시세']
            
            # 등락률(%) 컬럼 인덱스 찾기
            col_list = list(final_df.columns)
            try:
                ratio_idx = col_list.index('등락률(%)') + 1
            except:
                ratio_idx = None

            for row in range(2, ws.max_row + 1):
                if ratio_idx:
                    val = ws.cell(row=row, column=ratio_idx).value
                    try:
                        ratio_num = float(val)
                        name_cell = ws.cell(row=row, column=2)
                        if ratio_num > 0:
                            name_cell.fill = fill_red
                            name_cell.font = white_font
                        elif ratio_num < 0:
                            name_cell.fill = fill_blue
                            name_cell.font = white_font
                    except: pass

                for col in range(1, len(col_list) + 1):
                    ws.cell(row=row, column=col).alignment = Alignment(horizontal='center')
                    # 숫자 포맷 (소수점 2자리)
                    if isinstance(ws.cell(row=row, column=col).value, (int, float)):
                        ws.cell(row=row, column=col).number_format = '#,##0.00'

            # 열 너비 조절
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 30
            ws.column_dimensions['C'].width = 25
            for i in range(4, 9):
                ws.column_dimensions[chr(64+i)].width = 15

        # 4. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {target_date_str} 미국 전 종목 시세 리포트\n나스닥/뉴욕거래소 전 종목의 시가, 종가, 등락률 정보를 포함하고 있습니다."
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] {len(final_df)}개 종목 전송 완료 ---")

    except Exception as e:
        import traceback
        print(f"오류 발생:\n{traceback.format_exc()}")

if __name__ == "__main__":
    asyncio.run(send_us_all_stocks_report())
