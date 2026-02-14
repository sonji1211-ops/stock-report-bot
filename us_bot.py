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

# 나스닥 100 주요 종목 한글 매핑 (필요한 것만 유지 가능)
KOR_NAMES = {'AAPL': '애플', 'MSFT': '마이크로소프트', 'NVDA': '엔비디아', 'AMZN': '아마존', 'TSLA': '테슬라', 'META': '메타', 'GOOGL': '알파벳A'}

async def send_us_nasdaq100_full_report():
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    if day_of_week == 6: report_type = "주간(월-금평균)"
    elif day_of_week == 5: report_type = "일일(금요일마감)"
    else: report_type = "일일"

    try:
        # 나스닥 상위 100개 데이터 수집
        df_nas = fdr.StockListing('NASDAQ').head(100)
        report_list = []

        for idx, row in df_nas.iterrows():
            ticker = row['Symbol']
            try:
                # 1. 데이터 수집 및 소수점 2자리 반올림 (round 사용)
                df_p = fdr.DataReader(ticker).tail(2)
                if len(df_p) < 2: continue
                curr, prev = df_p.iloc[-1], df_p.iloc[-2]
                
                chg = ((curr['Close'] - prev['Close']) / prev['Close']) * 100
                
                # 리스트에 담을 때 미리 반올림하여 긴 소수점 차단
                report_list.append({
                    '티커': ticker, 
                    '종목명': KOR_NAMES.get(ticker, row['Name']), 
                    '시가($)': round(curr['Open'], 2), 
                    '고가($)': round(curr['High'], 2), 
                    '저가($)': round(curr['Low'], 2), 
                    '종가($)': round(curr['Close'], 2), 
                    '등락률(%)': round(chg, 2)
                })
            except: continue

        if not report_list: return
        
        # 2. 데이터프레임 변환 및 정렬
        df_final = pd.DataFrame(report_list).sort_values(by='등락률(%)', ascending=False)
        file_name = f"{now.strftime('%Y-%m-%d')}_{report_type}_미국리포트.xlsx"

        # 3. 엑셀 파일 생성 및 색상 분리
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            
            for row in range(2, ws.max_row + 1):
                # [중요] 미국장은 등락률이 7번째 열입니다.
                val = abs(float(ws.cell(row=row, column=7).value or 0))
                name_cell = ws.cell(row=row, column=2) # 종목명 칸
                
                # 지수님 요청 4단계 색상 필터
                if val >= 25:
                    name_cell.fill, name_cell.font = fill_red, font_white
                elif val >= 20:
                    name_cell.fill = fill_orange
                elif val >= 10:
                    name_cell.fill = fill_yellow
                
                # 4. 엑셀 표시 형식 최적화 (가운데 정렬 + 소수점 2자리 강제)
                for col in range(1, 8):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    if col >= 3: # 시가, 고가, 저가, 종가, 등락률
                        cell.number_format = '0.00'
            
            ws.column_dimensions['B'].width = 25 # 종목명 너비

        # 5. 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {now.strftime('%Y-%m-%d')} {report_type} 나스닥 리포트\n✅ 소수점 2자리 고정 & 색상 로직 적용"
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
    except Exception as e: print(e)

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_full_report())
