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

# [통합 자산 리스트] 데이터 안정성을 위해 티커 포맷 최적화
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 'USD/KRW': '달러/원 환율',
    
    # 국내 ETF (안정성을 위해 .KS 접미사 사용 고려)
    '069500': 'KODEX 200', '252670': 'KODEX 200선물인버스2X', '305720': 'KODEX 2차전지산업',
    '455810': 'TIGER 미국배당다우존스', '462330': 'KODEX AI반도체핵심공정', '122630': 'KODEX 레버리지',
    
    # 가상화폐 (야후 파이낸스 표준 포맷)
    'BTC-KRW': '비트코인', 'ETH-KRW': '이더리움', 'XRP-KRW': '리플(XRP)', 
    'SOL-KRW': '솔라나(SOL)', 'USDT-KRW': '테더(USDT)',
    
    # 미국 ETF (40종 전체 복구)
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
    'GC=F': '금 선물', 'SI=F': '은 선물'
}

async def fetch_asset_data(symbol, start, end, mode):
    try:
        # 데이터 수집 시도
        df = fdr.DataReader(symbol, start, end)
        
        # 만약 비어있다면 국내 종목의 경우 소스 변경 시도
        if (df is None or df.empty) and symbol.isdigit():
             df = fdr.DataReader(f"KRX:{symbol}", start, end)

        if df is None or df.empty or len(df) < 2:
            return None
        
        last_close = float(df.iloc[-1]['Close'])
        prev_close = float(df.iloc[-2]['Close'])
        
        ratio = round(((last_close - prev_close) / prev_close) * 100, 2)
        if mode != 'daily':
            first_open = float(df.iloc[0]['Open'])
            ratio = round(((last_close - first_open) / first_open) * 100, 2)
            
        return {
            '티커': symbol, 
            '항목명': ASSET_NAMES.get(symbol, symbol), 
            '현재가': last_close, 
            '등락률': ratio, 
            '기준일': df.index[-1].strftime('%Y-%m-%d')
        }
    except:
        return None

async def send_etf_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    # 데이터 확보를 위해 조회 기간을 넉넉히 설정
    s_date = (now - timedelta(days=20)).strftime('%Y-%m-%d')
    e_date = now.strftime('%Y-%m-%d')
    mode = 'weekly' if now.weekday() == 6 else 'daily'

    tasks = [fetch_asset_data(s, s_date, e_date, mode) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df = pd.DataFrame([r for r in results if r is not None])
    
    if df.empty: return

    file_name = f"{now.strftime('%m%d')}_종합_자산_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        # 정렬: 티커 순서대로 유지
        df[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        # 스타일 설정
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 15
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, 5):
                cell = ws.cell(row, col)
                # 항목명(B) 왼쪽, 나머지 중앙
                cell.alignment = Alignment(horizontal='center', vertical='center') if col != 2 else Alignment(horizontal='left', vertical='center', indent=1)
                
                if row > 1:
                    # 숫자 포맷 (원화는 콤마만, 달러는 소수점까지)
                    t = str(ws.cell(row, 1).value)
                    cell.number_format = '#,##0' if ('-KRW' in t or t.isdigit() or t in ['KS11','KQ11']) else '#,##0.00'
                    if col == 4: cell.number_format = '0.00'
                    
                    # 3% 강조
                    if col == 2 and abs(float(ws.cell(row, 4).value or 0)) >= 3:
                        cell.fill = PatternFill("solid", fgColor="FFFF00")
                        cell.font = Font(bold=True)

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"🌍 종합 자산 리포트 ({now.strftime('%Y-%m-%d')})")

if __name__ == "__main__":
    asyncio.run(send_etf_report())