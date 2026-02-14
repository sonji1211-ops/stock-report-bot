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

# 나스닥 100 주요 종목 한글 매핑
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

    # 요일별 리포트 성격 설정
    if day_of_week == 6:
        report_type = "주간(월-금평균)"
    elif day_of_week == 5:
        report_type = "일일(금요일마감)"
    else:
        report_type = "일일"

    try:
        print(f"--- 미국 나스닥 100 {report_type} 분석 시작 ---")
        
        df_nas = fdr.StockListing('NASDAQ')
        top_100_tickers = df_nas.head(100)
        report_list = []

        for idx, row in top_100_tickers.iterrows():
            ticker = row['Symbol']
            name = KOR_NAMES.get(ticker, row['Name']) 
            
            try:
                # 데이터 수집
                df_price = fdr.DataReader(ticker).tail(2)
                if len(df_price) < 2: continue
                
                curr = df_price.iloc[-1]
                prev = df_price.iloc[-2]
                chg_ratio = ((curr['Close'] - prev['Close']) / prev['Close']) * 100
                
                # 소수점 2자리 반올림 최적화
                report_list.append({
                    '티커': ticker, '종목명': name, 
                    '시가($)': round(curr['Open'], 2), 
                    '고가($)': round(curr['High'], 2), 
                    '저가($)': round(curr['Low'], 2), 
                    '종가($)': round(curr['Close'], 2), 
                    '등락률(%)': round(chg_ratio, 2)
                })
            except: continue

        if not report_list: return

        # 데이터프레임 변환 및 정렬
        df_final = pd.DataFrame(report_list).sort_values(by='등락률(%)', ascending=False)
        file_name = f"{target_date_str}_{report_type}_미국리포트.xlsx"

        # 스타일 설정
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='NASDAQ100', index=False)
            ws = writer.sheets['NASDAQ100']
            
            for row in range(2, ws.max_row + 1):
                # 등락률(%) 열 (7번째)
                ratio_cell = ws.cell(row=row, column=7)
                val = abs(float(ratio_cell.value or 0))
                name_cell = ws.cell(row=row, column=2)

                # 색상 기준 (지수님 요청 4단계)
                if val >= 25: 
                    name_cell.fill, name_cell.font = fill_red, font_white
                elif val >= 20: 
                    name_cell.fill = fill_orange
                elif val >= 10: 
                    name_cell.fill = fill_yellow
                
                # 정렬 및 표시 형식
                for col in range(1, 8):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center')
                    # 숫자는 엑셀에서도 소수점 2자리로 보이게 고정
                    if col >= 3:
                        cell.number_format = '0.00'
            
            ws.column_dimensions['B'].width = 28 

        bot = Bot(token=TOKEN)
        async with bot:
            cap = f"🇺🇸 {target_date_str} 나스닥100 {report_type} 리포트"
            msg = f"{cap}\n\n✅ 소수점 2자리 최적화 완료\n⚪ 5%↑ | 🟡 10%↑ | 🟠 20%↑ | 🔴 25%↑"
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_nasdaq100_full_report())
