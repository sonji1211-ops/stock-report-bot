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

# [주요 종목 한글 매핑]
KR_NAMES = {
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
    'ORCL': '오라클', 'CTSH': '코그니전트', 'TTD': '더 트레이드 Desk', 'ON': '온 세미컨덕터',
    'CEG': '컨스텔레이션 에너지', 'MDB': '몽고DB', 'ANSS': '앤시스', 'SPLK': '스플렁크',
    'FAST': '패스널', 'DASH': '도어대시', 'ZSC': '지스케일러', 'ILMN': '일루미나',
    'WBD': '워너 브라더스', 'AZN': '아스트라제네카', 'SGEN': '시애틀 제네틱스'
}

async def fetch_us_stock(row, search_start, search_end, mode):
    try:
        symbol = row['Symbol']
        h = fdr.DataReader(symbol, search_start, search_end)
        if h.empty or len(h) < 2: return None
        
        last_idx = h.index[-1]
        last_close = h.loc[last_idx, 'Close']
        
        if mode == 'daily':
            prev_idx = h.index[-2]
            prev_close = h.loc[prev_idx, 'Close']
            ratio = round(((last_close - prev_close) / prev_close) * 100, 2)
            final_date = last_idx.strftime('%Y-%m-%d')
        else:
            first_open = h.iloc[0]['Open']
            ratio = round(((last_close - first_open) / first_open) * 100, 2)
            final_date = f"{h.index[0].strftime('%m%d')}~{h.index[-1].strftime('%m%d')}"
        
        return {
            '티커': symbol,
            '종목명': KR_NAMES.get(symbol, row.get('Name', symbol)),
            '종가': last_close,
            '등락률': ratio,
            '산업': row.get('Industry', '-'),
            '기준일': final_date
        }
    except:
        return None

async def send_us_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    search_end = now.strftime('%Y-%m-%d')
    search_start = (now - timedelta(days=10)).strftime('%Y-%m-%d')
    mode = 'weekly' if day_of_week == 6 else 'daily'

    try:
        print(f"--- 분석 시작 (모드: {mode}) ---")
        df_base = fdr.StockListing('NASDAQ')
        df_target = df_base.head(800)

        tasks = [fetch_us_stock(row, search_start, search_end, mode) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        
        df_raw = pd.DataFrame([r for r in results if r is not None])
        
        if df_raw.empty:
            print("수집된 데이터가 없습니다.")
            return

        most_common_date = df_raw['기준일'].value_counts().idxmax()
        df_final = df_raw[df_raw['기준일'] == most_common_date]

        up_df = df_final[df_final['등락률'] >= 5].sort_values('등락률', ascending=False)
        down_df = df_final[df_final['등락률'] <= -5].sort_values('등락률', ascending=True)

        file_name = f"{now.strftime('%m%d')}_미국장_리포트.xlsx"
        h_map = {'티커':'티커', '종목명':'종목명', '종가':'종가', '등락률':'등락률(%)', '산업':'산업'}

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in [('상승_TOP', up_df), ('하락_TOP', down_df)]:
                data[['티커','종목명','종가','등락률','산업']].rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 가독성 개선 1: 컬럼 너비 최적화
                ws.column_dimensions['A'].width = 12  # 티커
                ws.column_dimensions['B'].width = 35  # 종목명 (충분히 넓게)
                ws.column_dimensions['C'].width = 15  # 종가
                ws.column_dimensions['D'].width = 15  # 등락률
                ws.column_dimensions['E'].width = 45  # 산업 (긴 텍스트 대비)

                for row in range(2, ws.max_row + 1):
                    ratio_val = abs(float(ws.cell(row, 4).value or 0))
                    name_cell = ws.cell(row, 2)
                    
                    # 가독성 개선 2: 종목명 강조 및 색상
                    if ratio_val >= 20: 
                        name_cell.fill = PatternFill("solid", fgColor="FFCC00")
                        name_cell.font = Font(bold=True)
                    elif ratio_val >= 10: 
                        name_cell.fill = PatternFill("solid", fgColor="FFFF00")
                    
                    ws.cell(row, 3).number_format = '#,##0.00'
                    ws.cell(row, 4).number_format = '0.00'
                    
                    # 가독성 개선 3: 정렬 조절
                    for c in range(1, 6):
                        if c == 2: # 종목명은 왼쪽 정렬이 가독성이 좋음
                            ws.cell(row, c).alignment = Alignment(horizontal='left', vertical='center', indent=1)
                        else:
                            ws.cell(row, c).alignment = Alignment(horizontal='center', vertical='center')

        async with bot:
            msg = (f"🇺🇸 미국 나스닥 리포트 ({most_common_date})\n\n"
                   f"📈 상승(5%↑): {len(up_df)}개\n"
                   f"📉 하락(5%↓): {len(down_df)}개\n\n"
                   f"💡 종목명 셀 너비 확장 및 가독성 최적화 완료")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
        print(f"전송 완료: {most_common_date}")

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_report())