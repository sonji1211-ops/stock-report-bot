import os
import FinanceDataReader as fdr
import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# [전종목 리스트] 위안화(CNY/KRW) 및 엔, 유로 포함 / 전체 종목 검수 완료
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 
    'USD/KRW': '달러/원 환율', 'JPY/KRW': '엔/원 환율', 'EUR/KRW': '유로/원 환율', 'CNY/KRW': '위안/원 환율',
    '069500': 'KODEX 200', '252670': 'KODEX 200선물인버스2X', '305720': 'KODEX 2차전지산업',
    '455810': 'TIGER 미국배당다우존스', '462330': 'KODEX AI반도체핵심공정', '122630': 'KODEX 레버리지',
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
    """국내(A) 및 미국(B) 주요 지수 비교 차트 생성"""
    start_d = (now - timedelta(days=30)).strftime('%Y-%m-%d')
    group_a = {'KS11': 'KOSPI', 'KQ11': 'KOSDAQ', 'USD/KRW': 'USD/KRW'}
    group_b = {'QQQ': 'NASDAQ 100', 'SPY': 'S&P 500', 'SOXX': 'Semiconductor'}

    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(11, 13))
    
    # 한국 & 환율
    for sym, name in group_a.items():
        df = fdr.DataReader(sym, start_d)
        if not df.empty:
            norm = (df['Close'] / df['Close'].iloc[0]) * 100
            ax1.plot(norm, label=name, marker='o', markersize=3)
    ax1.set_title('Domestic Indices & USD/KRW (Base 100)', fontsize=14)
    ax1.legend(); ax1.grid(True, linestyle='--')

    # 미국 지수
    for sym, name in group_b.items():
        df = fdr.DataReader(sym, start_d)
        if not df.empty:
            norm = (df['Close'] / df['Close'].iloc[0]) * 100
            ax2.plot(norm, label=name, marker='s', markersize=3)
    ax2.set_title('US Major Indices (Base 100)', fontsize=14)
    ax2.legend(); ax2.grid(True, linestyle='--')

    chart_file = "market_summary.png"
    plt.tight_layout()
    plt.savefig(chart_file)
    plt.close()
    await bot.send_photo(CHAT_ID, open(chart_file, 'rb'), caption=f"📈 지수 추이 요약 ({now.strftime('%m/%d')})\n상단: 국장&환율 / 하단: 미장 핵심지수")

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

    # 1. 시각화 차트 전송 (A+B 통합)
    await create_market_chart(bot, now)

    # 2. 상세 엑셀 리포트 수집
    tasks = [fetch_asset_data(s, s_date) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df = pd.DataFrame([r for r in results if r is not None])
    
    file_name = f"{now.strftime('%m%d')}_종합_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        # 셀 크기 고정 및 정렬
        ws.column_dimensions['A'].width = 16
        ws.column_dimensions['B'].width = 32
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 14
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, 5):
                cell = ws.cell(row, col)
                # 정렬: 항목명(B)만 왼쪽, 나머지는 전부 중앙
                cell.alignment = Alignment(horizontal='center', vertical='center') if col != 2 else Alignment(horizontal='left', vertical='center', indent=1)
                
                if row > 1:
                    t = str(ws.cell(row, 1).value)
                    # ₩ 기호 자동 적용 (코인, 국주, 지수, KRW환율)
                    if '-KRW' in t or t.isdigit() or t in ['KS11', 'KQ11'] or '/KRW' in t:
                        ws.cell(row, 3).number_format = '"₩"#,##0.00'
                    else:
                        ws.cell(row, 3).number_format = '#,##0.00'

    await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"📊 전종목 상세 리포트 송부 완료\n(위안화/엔/유로 환율 및 455810 포함)")

if __name__ == "__main__":
    asyncio.run(send_total_report())