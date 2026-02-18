import os
import FinanceDataReader as fdr
import pandas as pd
import matplotlib
# GUI 없는 환경에서도 차트 생성이 가능하도록 설정
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# [전종목 리스트] 455810 제외 / 환율 3종 포함
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 
    'USD/KRW': '달러/원 환율', 'JPY/KRW': '엔/원 환율', 'EUR/KRW': '유로/원 환율', 'CNY/KRW': '위안/원 환율',
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
    'XLF': '금융섹터', 'XLV': '헬스케어섹터', 'XLE': '에너지섹터', 'XLK': '기술주섹터', 
    'XLY': '임의소비재', 'XLP': '필수소비재', 'GDX': '금광업', 'GLD': '금선물',
    'VNQ': '리츠(부동산)', 'BITO': '비트코인ETF', 'FNGU': '빅테크플러스(3배)', 'BULZ': '빅테크성장(3배)',
    'VTI': '미국전체주식', 'VXUS': '미국외전세계', 'VT': '전세계주식',
    'GC=F': '금 선물', 'SI=F': '은 선물'
}

async def create_market_chart(bot, now):
    """지수 수익률 추이 차트 생성 (A+B안 통합)"""
    try:
        start_d = (now - timedelta(days=30)).strftime('%Y-%m-%d')
        group_a = {'KS11': 'KOSPI', 'KQ11': 'KOSDAQ', 'USD/KRW': 'USD/KRW'}
        group_b = {'QQQ': 'NASDAQ100', 'SPY': 'S&P500', 'SOXX': 'SOXX'}

        plt.style.use('bmh')
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 12))
        
        for sym, name in group_a.items():
            df = fdr.DataReader(sym, start_d)
            if not df.empty:
                norm = (df['Close'] / df['Close'].iloc[0]) * 100
                ax1.plot(norm, label=name, linewidth=2)
        ax1.set_title('Domestic & FX Trends (Base 100)')
        ax1.legend(); ax1.grid(True)

        for sym, name in group_b.items():
            df = fdr.DataReader(sym, start_d)
            if not df.empty:
                norm = (df['Close'] / df['Close'].iloc[0]) * 100
                ax2.plot(norm, label=name, linewidth=2)
        ax2.set_title('US Market Trends (Base 100)')
        ax2.legend(); ax2.grid(True)

        chart_path = "summary_chart.png"
        plt.tight_layout()
        plt.savefig(chart_path)
        plt.close()
        await bot.send_photo(CHAT_ID, open(chart_path, 'rb'), caption=f"📊 Market Summary ({now.strftime('%m/%d')})")
    except Exception as e:
        print(f"차트 생성 실패: {e}")

async def fetch_asset_data(symbol, s_date):
    try:
        df = fdr.DataReader(symbol, s_date)
        if df is None or df.empty or len(df) < 2: return None
        last_c, prev_c = float(df.iloc[-1]['Close']), float(df.iloc[-2]['Close'])
        ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
        return {'티커': symbol, '항목명': ASSET_NAMES.get(symbol, symbol), '현재가': last_c, '등락률': ratio}
    except: return None

async def send_total_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    s_date = (now - timedelta(days=30)).strftime('%Y-%m-%d')

    # 1. 차트 생성 및 전송
    await create_market_chart(bot, now)

    # 2. 엑셀 데이터 수집
    tasks = [fetch_asset_data(s, s_date) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df = pd.DataFrame([r for r in results if r is not None])
    
    # 3. 엑셀 파일 생성
    file_name = f"{now.strftime('%m%d')}_종합_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['B'].width = 32
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 14
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, 5):
                cell = ws.cell(row, col)
                # 정렬 설정
                cell.alignment = Alignment(horizontal='center', vertical='center') if col != 2 else Alignment(horizontal='left', indent=1)
                
                if row > 1 and col == 3:
                    t = str(ws.cell(row, 1).value)
                    if '-KRW' in t or t.isdigit() or t in ['KS11', 'KQ11'] or '/KRW' in t:
                        cell.number_format = '"₩"#,##0.00'
                    else:
                        cell.number_format = '#,##0.00'

    # [수정] 요청하신 캡션 문구로 변경
    await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"🌍 종합 자산 리포트 ({now.strftime('%Y-%m-%d')})\n✅ 환율 및 주요 ETF")

if __name__ == "__main__":
    asyncio.run(send_total_report())