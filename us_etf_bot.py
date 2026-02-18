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

# [통합 자산 리스트]
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수',
    'BTC/USD': '비트코인', 'ETH/USD': '이더리움',
    'GC=F': '금 선물', 'SI=F': '은 선물', 'USD/KRW': '달러/원 환율',
    'QQQ': '나스닥100', 'TQQQ': '나스닥100(3배)', 'SQQQ': '나스닥100인버스(3배)',
    'SPY': 'S&P500', 'IVV': 'S&P500(iShares)', 'VOO': 'S&P500(Vanguard)',
    'DIA': '다우존스', 'IWM': '러셀2000', 'SOXX': '필라델피아반도체', 'SOXL': '반도체강세(3배)',
    'SOXS': '반도체약세(3배)', 'NVDL': '엔비디아(2배)', 'TSLL': '테슬라(2배)',
    'SCHD': '슈드(배당성장)', 'JEPI': '제피(고배당)', 'TLT': '미국채20년(장기채)',
    'TMF': '장기채강세(3배)', 'TMV': '장기채약세(3배)', 'ARKK': '아크혁신(캐시우드)',
    'XLF': '금융섹터', 'XLV': '헬스케어섹터', 'XLE': '에너지섹터', 'XLK': '기술주섹터',
    'XLY': '임의소비재', 'XLP': '필수소비재', 'GDX': '금광업', 'GLD': '금선물',
    'VNQ': '리츠(부동산)', 'BITO': '비트코인ETF', 'CONL': '코인베이스(2배)',
    'QLD': '나스닥100(2배)', 'SSO': 'S&P500(2배)', 'Upro': 'S&P500(3배)',
    'FNGU': '빅테크플러스(3배)', 'BULZ': '빅테크성장(3배)', 'SMH': '반도체ETF(VanEck)',
    'VTI': '미국전체주식', 'VXUS': '미국외전세계', 'VT': '전세계주식'
}

async def fetch_asset_data(symbol, search_start, search_end, mode):
    try:
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
            
        return {'티커': symbol, '항목명': ASSET_NAMES.get(symbol, symbol), '현재가': last_close, '등락률': ratio, '기준일': final_date}
    except: return None

async def send_etf_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()
    search_end = now.strftime('%Y-%m-%d')
    search_start = (now - timedelta(days=15)).strftime('%Y-%m-%d')
    mode = 'weekly' if day_of_week == 6 else 'daily'

    tasks = [fetch_asset_data(s, search_start, search_end, mode) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df_raw = pd.DataFrame([r for r in results if r is not None])
    if df_raw.empty: return

    most_common_date = df_raw['기준일'].value_counts().idxmax()
    df_final = df_raw[df_raw['기준일'] == most_common_date].sort_values('등락률', ascending=False)

    file_name = f"{now.strftime('%m%d')}_종합_자산_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_final[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='종합현황', index=False)
        ws = writer.sheets['종합현황']
        
        # 1. 컬럼 너비 설정
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 18 # 현재가
        ws.column_dimensions['D'].width = 15 # 등락률
        
        # 2. 스타일 및 정렬 적용
        for row in range(1, ws.max_row + 1): # 헤더 포함 정렬
            for col in range(1, 5):
                cell = ws.cell(row, col)
                
                # 정렬 규칙 적용
                if col == 2: # 항목명 (B열) - 왼쪽 정렬
                    cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
                else: # 티커(A), 현재가(C), 등락률(D) - 중앙 정렬
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 데이터 행 스타일 (2행부터)
                if row > 1:
                    if col == 4: # 등락률 데이터 포맷
                        cell.number_format = '0.00'
                    if col == 3: # 현재가 데이터 포맷
                        cell.number_format = '#,##0.00'
                    
                    # 3% 이상 변동 시 강조 (항목명 셀)
                    ratio = float(ws.cell(row, 4).value or 0)
                    if col == 2 and abs(ratio) >= 3:
                        cell.fill = PatternFill("solid", fgColor="FFFF00")
                        cell.font = Font(bold=True)

    async with bot:
        title = "🗓 [주간]" if mode == 'weekly' else "🌍 [종합]"
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"{title} 자산 종합 리포트 ({most_common_date})")

if __name__ == "__main__":
    asyncio.run(send_etf_report())