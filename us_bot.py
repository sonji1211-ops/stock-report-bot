import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

# 주요 100대 종목 한글 매핑 사전 (자주 보시는 종목 위주)
KOR_NAMES = {
    'AAPL': '애플', 'MSFT': '마이크로소프트', 'NVDA': '엔비디아', 'AMZN': '아마존',
    'GOOGL': '알파벳A', 'GOOG': '알파벳C', 'META': '메타(페이스북)', 'TSLA': '테슬라',
    'AVGO': '브로드컴', 'PEP': '펩시코', 'COST': '코스트코', 'ADBE': '어도비',
    'CSCO': '시스코 시스템즈', 'NFLX': '넷플릭스', 'AMD': 'AMD', 'TMUS': '티모바일',
    'INTU': '인튜이트', 'INTC': '인텔', 'AMAT': '어플라이드 머티어리얼즈', 'QCOM': '퀄컴',
    'TXN': '텍사스 인스트루먼트', 'AMGN': '암젠', 'ISRG': '인튜이티브 서지컬', 'HON': '허니웰',
    'BKNG': '부킹홀딩스', 'VRTX': '버텍스 파마슈티컬스', 'GILD': '길리어드 사이언스',
    'SBUX': '스타벅스', 'MDLZ': '몬델리즈', 'ADP': 'ADP', 'PANW': '팔로알토 네트웍스',
    'MELI': '메르카도리브레', 'REGN': '리제네론', 'MU': '마이크론 테크놀로지', 'SNPS': '시놉시스',
    'KLAC': 'KLA', 'CDNS': '케이던스 디자인', 'PYPL': '페이팔', 'MAR': '메리어트',
    'ASML': 'ASML', 'LRCX': '램 리서치', 'MNST': '몬스터 베버리지', 'ORLY': '오라일리',
    'ADSK': '오토데스크', 'LULU': '룰루레몬', 'KDP': '큐리그 닥터 페퍼', 'PAYX': '페이첵스',
    'FTNT': '포티넷', 'CHTR': '차터 커뮤니케이션즈', 'AEP': '아메리칸 일렉트릭 파워'
    # ... 필요한 종목은 여기에 계속 추가 가능합니다.
}

async def send_us_nasdaq100_korean_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    try:
        print(f"--- 나스닥 100 한글 리포트 시작 ---")
        
        df_nas = fdr.StockListing('NASDAQ')
        top_100_tickers = df_nas.head(100)

        report_list = []

        for idx, row in top_100_tickers.iterrows():
            ticker = row['Symbol']
            # 한글 이름 사전에 있으면 한글로, 없으면 영문명 사용
            name = KOR_NAMES.get(ticker, row['Name']) 
            
            try:
                df = fdr.DataReader(ticker).tail(2)
                if len(df) < 2: continue
                
                prev_close = df.iloc[0]['Close']
                curr_close = df.iloc[1]['Close']
                curr_open = df.iloc[1]['Open']
                
                chg_ratio = ((curr_close - prev_close) / prev_close) * 100

                report_list.append({
                    '티커': ticker,
                    '종목명': name,
                    '시작가($)': curr_open,
                    '마감가($)': curr_close,
                    '등락률(%)': chg_ratio
                })
            except:
                continue

        df_final = pd.DataFrame(report_list)
        file_name = f"{target_date_str}_나스닥100_한글리포트.xlsx"
        
        font_red = Font(color="FF0000", bold=True)
        font_blue = Font(color="0000FF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            
            for row in range(2, ws.max_row + 1):
                ratio_val = ws.cell(row=row, column=5).value
                name_cell = ws.cell(row=row, column=2)
                ratio_cell = ws.cell(row=row, column=5)
                
                if ratio_val is not None:
                    if ratio_val > 0:
                        name_cell.font = font_red
                        ratio_cell.font = font_red
                    elif ratio_val < 0:
                        name_cell.font = font_blue
                        ratio_cell.font = font_blue

                for col in range(1, 6):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

            ws.column_dimensions['B'].width = 25
            for i in range(3, 6):
                ws.column_dimensions[chr(64+i)].width = 15

        bot = Bot(token=TOKEN)
        async with bot:
            msg = f"🇺🇸 {target_date_str} 나스닥 100 한글 리포트\n주요 종목이 한글로 표시되어 보기 편합니다!"
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] 리포트 전송 완료 ---")

    except Exception as e:
        print(f"에러 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_korean_report())
