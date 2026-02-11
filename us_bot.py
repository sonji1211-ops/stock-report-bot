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

async def send_us_major_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    try:
        print(f"--- 미국 주요 종목 분석 시작: {target_date_str} ---")
        
        # 1. 나스닥 100 종목 리스트 가져오기
        # NASDAQ 100은 나스닥의 핵심 우량주 100개를 의미합니다.
        df_nas100 = fdr.StockListing('NASDAQ')
        
        # 시가총액 순으로 상위 100개만 자릅니다 (애플, 마이크로소프트, 엔비디아 등 포함)
        df_top100 = df_nas100.head(100).copy()

        # 2. 한글 매핑 및 정리
        # 미국 데이터는 컬럼명이 다를 수 있어 유연하게 매핑합니다.
        h_map = {
            'Symbol': '티커(코드)', 
            'Name': '종목명', 
            'Industry': '산업군',
            'Price': '현재가($)', 
            'Changes': '전일대비', 
            'ChgPct': '등락률(%)'
        }
        
        # 실제 존재하는 컬럼만 선택
        df_final = df_top100[[c for c in h_map.keys() if c in df_top100.columns]].copy()
        df_final = df_final.rename(columns=h_map)

        # 3. 엑셀 파일 생성 및 스타일 적용
        file_name = f"{target_date_str}_나스닥100_리포트.xlsx"
        
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # 상승
        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid") # 하락
        white_font = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='나스닥상위100', index=False)
            ws = writer.sheets['나스닥상위100']
            
            # 등락률(%) 컬럼 위치 찾기 (보통 마지막)
            ratio_idx = len(df_final.columns)
            
            for row in range(2, ws.max_row + 1):
                ratio_val = ws.cell(row=row, column=ratio_idx).value
                name_cell = ws.cell(row=row, column=2) # 종목명 칸 색칠
                
                try:
                    ratio_num = float(ratio_val)
                    if ratio_num > 0:
                        name_cell.fill = fill_red
                        name_cell.font = white_font
                    elif ratio_num < 0:
                        name_cell.fill = fill_blue
                        name_cell.font = white_font
                except:
                    pass

                # 정렬 및 서식
                for col in range(1, len(df_final.columns) + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

            # 열 너비 조절
            for i in range(1, len(df_final.columns) + 1):
                ws.column_dimensions[chr(64+i)].width = 20

        # 4. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {target_date_str} 나스닥 100 주요 종목 리포트\n시가총액 상위 100개 종목의 마감 현황입니다."
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] 미국 종목 리포트 전송 완료 ---")

    except Exception as e:
        import traceback
        print(f"오류 상세:\n{traceback.format_exc()}")

if __name__ == "__main__":
    asyncio.run(send_us_major_report())
