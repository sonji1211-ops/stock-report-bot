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

# [전종목 리스트]
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 
    'USD/KRW': '달러/원 환율', 'JPY/KRW': '엔/원 환율', 
    'EUR/KRW': '유로/원 환율', 'CNY/KRW': '위안/원 환율', 
    '069500': 'KODEX 200', '252670': 'KODEX 200선물인버스2X', '305720': 'KODEX 2차전지산업',
    '462330': 'KODEX AI반도체핵심공정', '122630': 'KODEX 레버리지',
    'BTC-KRW': '비트코인', 'ETH-KRW': '이더리움', 'XRP-KRW': '리플(XRP)', 
    'SOL-KRW': '솔라나(SOL)', 'USDT-KRW': '테더(USDT)',
    'QQQ': '나스닥100', 'TQQQ': '나스닥100(3배)', 'SQQQ': '나스닥100인버스(3배)', 'QLD': '나스닥100(2배)',
    'SPY': 'S&P500', 'IVV': 'S&P500(iShares)', 'VOO': 'S&P500(Vanguard)', 'SSO': 'S&P500(2배)', 'Upro': 'S&P500(3배)',
    'DIA': '다우존스', 'IWM': '러셀2000', 'SOXX': '필라델피아반도체', 'SOXL': '반도체강세(3배)', 'SOXS': '반도체약세(3배)',
    'SMH': '반도체ETF(VanEck)', 'NVDL': '엔비디아(2배)', 'TSLL': '테슬라(2배)', 'CONL': '코인베이스(2배)',
    'SCHD': '슈드(배당성장)', 'JEPI': '제피(고배당)', 'ARKK': '아크혁신(캐시우드)',
    'TLT': '미국채20년(장기채)', 'TMF': '장기채강세(3배)', 'TMV': '장기채약세(3배)',
    'XLF': '금융섹터', 'XLV': '헬스케어섹터', 'XLE': 'energy섹터', 'XLK': '기술주섹터', 
    'XLY': '임의소비재', 'XLP': '필수소비재', 'GDX': '금광업', 'GLD': '금선물',
    'VNQ': '리츠(부동산)', 'BITO': '비트코인ETF', 'FNGU': '빅테크플러스(3배)', 'BULZ': '빅테크성장(3배)',
    'VTI': '미국전체주식', 'VXUS': '미국외전세계', 'VT': '전세계주식',
    'GC=F': '금 선물', 'SI=F': '은 선물'
}

async def fetch_asset_data(symbol, s_date):
    try:
        df = fdr.DataReader(symbol, s_date)
        
        # 위안화 누락 방지: 첫 번째 티커 실패 시 대체 티커 시도
        if (df is None or df.empty) and symbol == 'CNY/KRW':
            df = fdr.DataReader('CNYKRW=X', s_date)
            
        if df is None or df.empty or len(df) < 2: return None
        
        last_c = float(df.iloc[-1]['Close'])
        prev_c = float(df.iloc[-2]['Close'])
        ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
            
        return {'티커': symbol, '항목명': ASSET_NAMES.get(symbol, symbol), '현재가': last_c, '등락률': ratio}
    except:
        return None

async def send_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    s_date = (now - timedelta(days=30)).strftime('%Y-%m-%d')

    tasks = [fetch_asset_data(s, s_date) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df = pd.DataFrame([r for r in results if r is not None])
    
    if df.empty: return

    file_name = f"{now.strftime('%m%d')}_종합_리포트.xlsx"
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 25
        ws.column_dimensions['D'].width = 15

        for row in range(1, ws.max_row + 1):
            # [수정] 모든 셀 가운데 정렬 (종목명 포함)
            for col in range(1, 5):
                ws.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')

            if row > 1:
                # [에러 방지] 빈 값 검사 후 숫자로 변환
                val = ws.cell(row, 4).value
                ratio_val = abs(float(val)) if val is not None and val != '' else 0

                # 3% 이상 강조
                if ratio_val >= 3:
                    for col in range(1, 5):
                        ws.cell(row, col).fill = yellow_fill
                        ws.cell(row, col).font = Font(bold=True)
                
                # 원화 표시 대상
                t = str(ws.cell(row, 1).value)
                if col == 4: # 루프 내 마지막 컬럼 처리 시 서식 적용
                    if '-KRW' in t or t.isdigit() or '/KRW' in t or 'KS11' in t:
                        ws.cell(row, 3).number_format = '"₩"#,##0.00'
                    else:
                        ws.cell(row, 3).number_format = '#,##0.00'
                    ws.cell(row, 4).number_format = '0.00'

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), 
                               caption=f"🌍 종합 리포트 ({now.strftime('%Y-%m-%d')})\n✅ 위안화 보강 및 전 항목 가운데 정렬 완료")

if __name__ == "__main__":
    asyncio.run(send_report())