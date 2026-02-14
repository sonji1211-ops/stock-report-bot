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
    'ORCL': '오라클', 'CTSH': '코그니전트', 'TTD': '더 트레이드 데스크', 'ON': '온 세미컨덕터',
    'CEG': '컨스텔레이션 에너지', 'MDB': '몽고DB', 'ANSS': '앤시스', 'SPLK': '스플렁크',
    'FAST': '패스널', 'DASH': '도어대시', 'ZSC': '지스케일러', 'ILMN': '일루미나',
    'WBD': '워너 브라더스', 'AZN': '아스트라제네카', 'SGEN': '시애틀 제네틱스'
}

async def send_us_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = (now - timedelta(days=1)).strftime('%Y-%m-%d')

    try:
        print("--- 나스닥 데이터 수집 중 ---")
        df_base = fdr.StockListing('NASDAQ')
        if df_base is None or df_base.empty: return

        # 수치형 변환 (오류 방지)
        df_base['Close'] = pd.to_numeric(df_base['Close'], errors='coerce').fillna(0)
        
        # 등락률 계산 (ChgPct가 있으면 사용, 없으면 직접 계산)
        if 'ChgPct' in df_base.columns:
            df_base['Ratio'] = pd.to_numeric(df_base['ChgPct'], errors='coerce').fillna(0) * 100
        else:
            # 직접 계산 시 'Close'와 'Changes' 컬럼 활용
            df_base['Changes'] = pd.to_numeric(df_base.get('Changes', 0), errors='coerce').fillna(0)
            df_base['Ratio'] = (df_base['Changes'] / (df_base['Close'] - df_base['Changes']) * 100).fillna(0)

        # 한글 이름 적용
        df_base['Name'] = df_base.apply(lambda x: KR_NAMES.get(x['Symbol'], x['Name']), axis=1)

        # 컬럼 순서 설정 (티커, 종목명, 종가, 등락률, 산업군)
        df_final = df_base[['Symbol', 'Name', 'Close', 'Ratio', 'Industry']].copy()
        
        up_df = df_final[df_final['Ratio'] >= 5].sort_values('Ratio', ascending=False)
        down_df = df_final[df_final['Ratio'] <= -5].sort_values('Ratio', ascending=True)

        # 엑셀 파일 생성
        file_name = f"{now.strftime('%m%d')}_나스닥_리포트.xlsx"
        h_map = {'Symbol':'티커', 'Name':'종목명', 'Close':'종가', 'Ratio':'등락률(%)', 'Industry':'산업'}

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in [('나스닥_상승', up_df), ('나스닥_하락', down_df)]:
                data.rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                for row in range(2, ws.max_row + 1):
                    ratio_val = abs(float(ws.cell(row, 4).value or 0)) # 등락률 컬럼(D열)
                    name_cell = ws.cell(row, 2) # 종목명(B열)
                    
                    # 색상 강조
                    if ratio_val >= 20: name_cell.fill = PatternFill("solid", fgColor="FFCC00")
                    elif ratio_val >= 10: name_cell.fill = PatternFill("solid", fgColor="FFFF00")
                    
                    # 가독성: 종가 천 단위 콤마(C열), 등락률 소수점(D열)
                    ws.cell(row, 3).number_format = '#,##0.00'
                    ws.cell(row, 4).number_format = '0.00'
                    
                    for c in range(1, 6):
                        ws.cell(row, c).alignment = Alignment(horizontal='center')
                for i in range(1, 6): ws.column_dimensions[chr(64+i)].width = 20

        async with bot:
            msg = (f"🇺🇸 {target_date_str} 나스닥 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(up_df)}개\n"
                   f"📉 하락(5%↓): {len(down_df)}개\n\n"
                   f"💡 주요 100개 종목 한글화 & 가독성 강화 적용")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=msg)
        print(f"--- {file_name} 전송 완료 ---")

    except Exception as e:
        print(f"미국장 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_report())
