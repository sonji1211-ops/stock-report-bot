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

# 나스닥 100 주요 종목 전체 한글 매핑 (100개 근접)
KOR_NAMES = {
    'AAPL': '애플', 'MSFT': '마이크로소프트', 'NVDA': '엔비디아', 'AMZN': '아마존',
    'GOOGL': '알파벳A', 'GOOG': '알파벳C', 'META': '메타', 'TSLA': '테슬라',
    'AVGO': '브로드컴', 'PEP': '펩시코', 'COST': '코스트코', 'ADBE': '어도비',
    'CSCO': '시스코', 'NFLX': '넷플릭스', 'AMD': 'AMD', 'TMUS': '티모바일',
    'INTU': '인튜이트', 'INTC': '인텔', 'AMAT': '어플라이드 머티어리얼즈', 'QCOM': '퀄컴',
    'TXN': '텍사스 인스트루먼트', 'AMGN': '암젠', 'ISRG': '인튜이티브 서지컬', 'HON': '허니웰',
    'BKNG': '부킹홀딩스', 'VRTX': '버텍스 파마슈티컬스', 'GILD': '길리어드 사이언스',
    'SBUX': '스타벅스', 'MDLZ': '몬델리즈', 'ADP': 'ADP', 'PANW': '팔로알토 네트웍스',
    'MELI': '메르카도리브레', 'REGN': '리제네론', 'MU': '마이크론 테크놀로지', 'SNPS': '시놉시스',
    'KLAC': 'KLA', 'CDNS': '케이던스 디자인', 'PYPL': '페이팔', 'MAR': '메리어트',
    'ASML': 'ASML', 'LRCX': '램 리서치', 'MNST': '몬스터 베버리지', 'ORLY': '오라일리',
    'ADSK': '오토데스크', 'LULU': '룰루레몬', 'KDP': '큐리그 닥터 페퍼', 'PAYX': '페이첵스',
    'FTNT': '포티넷', 'CHTR': '차터 커뮤니케이션즈', 'AEP': '아메리칸 일렉트릭 파워',
    'PDD': '핀둬둬', 'NXPI': 'NXP 세미컨덕터', 'DXCM': '덱스콤', 'MCHP': '마이크로칩',
    'CPRT': '코파트', 'ROST': '로스 스토어', 'IDXX': '아이덱스 래버러토리', 'PCAR': '파카',
    'CSX': 'CSX', 'ODFL': '올드 도미니언', 'KVUE': '켄뷰', 'EXC': '엑셀론',
    'BKR': '베이커 휴즈', 'GEHC': 'GE 헬스케어', 'CTAS': '신타스', 'WDAY': '워크데이',
    'TEAM': '아틀라시안', 'DDOG': '데이터독', 'MRVL': '마벨 테크놀로지', 'ABNB': '에어비앤비',
    'ORCL': '오라클', 'CTSH': '코그니전트', 'TTD': '더 트레이드 데스크', 'ON': '온 세미컨덕터',
    'CEG': '컨스텔레이션 에너지', 'MDB': '몽고DB', 'ANSS': '앤시스', 'SPLK': '스플렁크',
    'FAST': '패스널', 'DASH': '도어대시', 'ZSC': '지스케일러', 'ILMN': '일루미나',
    'WBD': '워너 브라더스', 'AZN': '아스트라제네카', 'SGEN': '시애틀 제네틱스'
}

async def send_us_nasdaq100_full_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')
    day_of_week = now.weekday() 

    is_saturday = (day_of_week == 5)

    try:
        print(f"--- 나스닥 100 한글 풀 리포트 시작 (토요일 필터링: {is_saturday}) ---")
        
        # 1. 나스닥 종목 리스팅
        df_nas = fdr.StockListing('NASDAQ')
        top_100_tickers = df_nas.head(100)

        report_list = []

        # 2. 각 종목별 데이터 수집
        for idx, row in top_100_tickers.iterrows():
            ticker = row['Symbol']
            # 매핑 사전에 있으면 한글, 없으면 원문 사용
            name = KOR_NAMES.get(ticker, row['Name']) 
            
            try:
                # 데이터 기간 확보
                df = fdr.DataReader(ticker).tail(7)
                if len(df) < 2: continue
                
                if is_saturday:
                    # 주간 누적 (월~금)
                    weekly_open = df.iloc[0]['Open']
                    weekly_close = df.iloc[-1]['Close']
                    chg_ratio = ((weekly_close - weekly_open) / weekly_open) * 100
                    
                    if abs(chg_ratio) >= 5:
                        report_list.append({
                            '티커': ticker, '종목명': name, '주초시작가($)': weekly_open,
                            '주말마감가($)': weekly_close, '주간등락률(%)': chg_ratio
                        })
                else:
                    # 일일 변동
                    prev_close = df.iloc[-2]['Close']
                    curr_close = df.iloc[-1]['Close']
                    curr_open = df.iloc[-1]['Open']
                    chg_ratio = ((curr_close - prev_close) / prev_close) * 100
                    
                    report_list.append({
                        '티커': ticker, '종목명': name, '시작가($)': curr_open,
                        '마감가($)': curr_close, '등락률(%)': chg_ratio
                    })
            except:
                continue

        if not report_list:
            if is_saturday:
                bot = Bot(token=TOKEN)
                async with bot:
                    await bot.send_message(chat_id=CHAT_ID, text=f"🇺🇸 {target_date_str}\n이번 주 5% 이상 변동 종목 없음")
                return
            else: return

        # 3. 엑셀 제작 및 전송
        df_final = pd.DataFrame(report_list)
        file_name = f"{target_date_str}_나스닥100_최종리포트.xlsx"
        font_red = Font(color="FF0000", bold=True)
        font_blue = Font(color="0000FF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            for row in range(2, ws.max_row + 1):
                ratio_val = ws.cell(row=row, column=5).value
                if ratio_val and ratio_val > 0:
                    ws.cell(row=row, column=2).font = font_red
                    ws.cell(row=row, column=5).font = font_red
                elif ratio_val and ratio_val < 0:
                    ws.cell(row=row, column=2).font = font_blue
                    ws.cell(row=row, column=5).font = font_blue
                for col in range(1, 6):
                    ws.cell(row=row, column=col).alignment = Alignment(horizontal='center')
            ws.column_dimensions['B'].width = 28

        bot = Bot(token=TOKEN)
        async with bot:
            cap = f"🇺🇸 {target_date_str} 나스닥 100 마감 리포트"
            if is_saturday: cap = f"📊 {target_date_str} 주간 누적(±5%) 리포트"
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=cap)

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_full_report())
