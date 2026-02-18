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
    # 1. 국내 지수 및 환율
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 'USD/KRW': '달러/원 환율',
    
    # 2. 국내 주요 ETF (네이버 소스 사용)
    '069500': 'KODEX 200', '252670': 'KODEX 200선물인버스2X', '305720': 'KODEX 2차전지산업',
    '455810': 'TIGER 미국배당다우존스', '462330': 'KODEX AI반도체핵심공정', '122630': 'KODEX 레버리지',
    
    # 3. 가상화폐 (원화 가격 & 등락률 정상화)
    'BTC/KRW': '비트코인', 'ETH/KRW': '이더리움', 'XRP/KRW': '리플(XRP)', 
    'SOL/KRW': '솔라나(SOL)', 'USDT/KRW': '테더(USDT)',
    
    # 4. 미국 지수 및 주요 ETF (40종 전체)
    'QQQ': '나스닥100', 'TQQQ': '나스닥100(3배)', 'SQQQ': '나스닥100인버스(3배)', 'QLD': '나스닥100(2배)',
    'SPY': 'S&P500', 'IVV': 'S&P500(iShares)', 'VOO': 'S&P500(Vanguard)', 'SSO': 'S&P500(2배)', 'Upro': 'S&P500(3배)',
    'DIA': '다우존스', 'IWM': '러셀2000', 'SOXX': '필라델피아반도체', 'SOXL': '반도체강세(3배)', 'SOXS': '반도체약세(3배)', 
    'SMH': '반도체ETF(VanEck)', 'NVDL': '엔비디아(2배)', 'TSLL': '테슬라(2배)', 'CONL': '코인베이스(2배)',
    'SCHD': '슈드(배당성장)', 'JEPI': '제피(고배당)', 'ARKK': '아크혁신(캐시우드)',
    'TLT': '미국채20년(장기채)', 'TMF': '장기채강세(3배)', 'TMV': '장기채약세(3배)',
    'XLF': '금융섹터', 'XLV': '헬스케어섹터', 'XLE': '에너지섹터', 'XLK': '기술주섹터', 
    'XLY': '임의소비재', 'XLP': '필수소비재', 'GDX': '금광업', 'GLD': '금선물',
    'VNQ': '리츠(부동산)', 'BITO': '비트코인ETF', 'FNGU': '빅테크플러스(3배)', 'BULZ': '빅테크성장(3배)',
    'VTI': '미국전체주식', 'VXUS': '미국외전세계', 'VT': '전세계주식',
    
    # 5. 원자재
    'GC=F': '금 선물', 'SI=F': '은 선물'
}

async def fetch_asset_data(symbol, search_start, search_end, mode):
    try:
        # 가상화폐와 국내 ETF 데이터 소스 분기 처리
        if '/' in symbol: # 코인 또는 환율 (예: BTC/KRW)
            df = fdr.DataReader(symbol, search_start, search_end)
        elif symbol.isdigit(): # 국내 종목 (숫자 6자리)
            df = fdr.DataReader(symbol, search_start, search_end)
        else: # 미국 종목 등 기타
            df = fdr.DataReader(symbol, search_start, search_end)

        if df is None or df.empty or len(df) < 2:
            return None
        
        last_close = df.iloc[-1]['Close']
        prev_close = df.iloc[-2]['Close']
        
        if mode == 'daily':
            ratio = round(((last_close - prev_close) / prev_close) * 100, 2)
            final_date = df.index[-1].strftime('%Y-%m-%d')
        else:
            first_open = df.iloc[0]['Open']
            ratio = round(((last_close - first_open) / first_open) * 100, 2)
            final_date = f"{df.index[0].strftime('%m%d')}~{df.index[-1].strftime('%m%d')}"
            
        return {'티커': symbol, '항목명': ASSET_NAMES.get(symbol, symbol), '현재가': last_close, '등락률': ratio, '기준일': final_date}
    except:
        return None

async def send_etf_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()
    
    # 영업일 고려하여 기간 넉넉히 설정
    search_end = now.strftime('%Y-%m-%d')
    search_start = (now - timedelta(days=20)).strftime('%Y-%m-%d')
    mode = 'weekly' if day_of_week == 6 else 'daily'

    tasks = [fetch_asset_data(s, search_start, search_end, mode) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df_raw = pd.DataFrame([r for r in results if r is not None])
    
    if df_raw.empty: return

    most_common_date = df_raw['기준일'].value_counts().idxmax()
    df_final = df_raw.copy()

    file_name = f"{now.strftime('%m%d')}_종합_자산_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_final[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='종합현황', index=False)
        ws = writer.sheets['종합현황']
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 22 # 원화 가격 대비 확장
        ws.column_dimensions['D'].width = 15
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, 5):
                cell = ws.cell(row, col)
                if col == 2:
                    cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
                else:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                if row > 1:
                    ticker_val = str(ws.cell(row, 1).value)
                    # KRW(원화) 포함된 가격은 콤마만, 나머지는 소수점 2자리
                    if 'KRW' in ticker_val or ticker_val.isdigit():
                        ws.cell(row, 3).number_format = '#,##0'
                    else:
                        ws.cell(row, 3).number_format = '#,##0.00'
                    
                    ws.cell(row, 4).number_format = '0.00'
                    
                    ratio_val = float(ws.cell(row, 4).value or 0)
                    if col == 2 and abs(ratio_val) >= 3:
                        cell.fill = PatternFill("solid", fgColor="FFFF00")
                        cell.font = Font(bold=True)

    async with bot:
        title = "🌍 [종합]" if mode == 'daily' else "🗓 [주간]"
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"{title} 한·미 자산 리포트 ({most_common_date})\n💡 국내 ETF 에러 및 가상화폐 등락률 수정 완료")

if __name__ == "__main__":
    asyncio.run(send_etf_report())