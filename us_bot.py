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

async def send_us_nasdaq100_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    try:
        print(f"--- 나스닥 100 상세 분석 시작: {target_date_str} ---")
        
        # 1. 나스닥 종목 리스팅 (시가총액 상위순으로 가져옴)
        df_nas = fdr.StockListing('NASDAQ')
        
        # 상위 100개 추출 (나스닥 100 주요 종목)
        df_top100 = df_nas.head(100).copy()

        # 2. 데이터 컬럼 정리 및 이름 변경
        # 미국 데이터 소스의 컬럼명을 한국식 리포트에 맞게 매핑합니다.
        # 소스에 따라 Price, Open, High, Low, ChangesRatio 등의 이름으로 들어옵니다.
        h_map = {
            'Symbol': '티커',
            'Name': '종목명',
            'Industry': '산업',
            'Price': '종가($)',
            'Open': '시가($)',
            'High': '고가($)',
            'Low': '저가($)',
            'ChangesRatio': '등락률(%)'
        }
        
        # 실제 존재하는 컬럼들만 골라서 리포트 생성
        cols_to_use = [c for c in h_map.keys() if c in df_top100.columns]
        df_final = df_top100[cols_to_use].copy()
        df_final = df_final.rename(columns=h_map)

        # 3. 엑셀 파일 생성 및 꾸미기
        file_name = f"{target_date_str}_나스닥100_시세리포트.xlsx"
        
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='나스닥100', index=False)
            ws = writer.sheets['나스닥100']
            
            # 등락률 컬럼 위치 확인
            col_names = list(df_final.columns)
            ratio_idx = col_names.index('등락률(%)') + 1 if '등락률(%)' in col_names else None

            for row in range(2, ws.max_row + 1):
                # 등락률에 따른 종목명 색상 입히기
                if ratio_idx:
                    val = ws.cell(row=row, column=ratio_idx).value
                    name_cell = ws.cell(row=row, column=2)
                    try:
                        ratio_num = float(val)
                        if ratio_num > 0:
                            name_cell.fill = fill_red
                            name_cell.font = white_font
                        elif ratio_num < 0:
                            name_cell.fill = fill_blue
                            name_cell.font = white_font
                    except: pass

                # 전체 셀 가운데 정렬 및 숫자 포맷
                for col in range(1, len(col_names) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

            # 열 너비 자동 조절
            ws.column_dimensions['A'].width = 12
            ws.column_dimensions['B'].width = 25
            ws.column_dimensions['C'].width = 25
            for i in range(4, 9):
                ws.column_dimensions[chr(64+i)].width = 15

        # 4. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {target_date_str} 나스닥 100 시세 리포트\n주요 100개 종목의 시가, 종가, 등락률 정보입니다."
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] 나스닥 100 리포트 전송 완료 ---")

    except Exception as e:
        import traceback
        print(f"오류 발생:\n{traceback.format_exc()}")

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_report())
