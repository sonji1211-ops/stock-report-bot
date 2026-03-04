import os, pandas as pd, asyncio, datetime
import yfinance as yf
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# [전종목 리스트]
ASSET_NAMES = {
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 
    'USD/KRW': '달러/원 환율', 'JPY/KRW': '엔/원 환율', 
    'EUR/KRW': '유로/원 환율', 'CNY/KRW': '위안/원 환율', 
    '069500.KS': 'KODEX 200', '252670.KS': 'KODEX 200선물인버스2X', '305720.KS': 'KODEX 2차전지산업',
    '462330.KS': 'KODEX AI반도체핵심공정', '122630.KS': 'KODEX 레버리지',
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

async def fetch_asset_data(symbol):
    try:
        yf_symbol = symbol
        if symbol == 'KS11': yf_symbol = '^KS11'
        elif symbol == 'KQ11': yf_symbol = '^KQ11'
        elif symbol == 'USD/KRW': yf_symbol = 'KRW=X'
        elif symbol == 'JPY/KRW': yf_symbol = 'JPYKRW=X'
        elif symbol == 'EUR/KRW': yf_symbol = 'EURKRW=X'
        elif symbol == 'CNY/KRW': yf_symbol = 'CNYKRW=X'
        elif symbol.isdigit(): yf_symbol = symbol + ".KS"

        ticker_obj = yf.Ticker(yf_symbol)
        df = ticker_obj.history(period="5d")
        df = df.dropna(subset=['Close'])
        
        if len(df) < 2: return None

        last_c = float(df['Close'].iloc[-1])
        prev_c = float(df['Close'].iloc[-2])
        
        # 등락률 자리에 들어갈 '전일 대비 숫자(수치)'
        diff_value = last_c - prev_c
            
        return {
            '티커': symbol, 
            '항목명': ASSET_NAMES.get(symbol, symbol), 
            '현재가': last_c, 
            '등락률%': diff_value # 이름은 등락률% 지만 수치값 전달
        }
    except: return None

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    results = []
    for s in ASSET_NAMES.keys():
        res = await fetch_asset_data(s)
        if res: results.append(res)
        await asyncio.sleep(0.05)

    df = pd.DataFrame(results)
    file_name = f"{now.strftime('%m%d')}_종합_리포트.xlsx"
    
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        header_fill = PatternFill(start_color='444444', end_color='444444', fill_type='solid')
        white_font = Font(color='FFFFFF', bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        # 가독성을 위한 칸 너비 조절
        ws.column_dimensions['A'].width = 15 # 티커
        ws.column_dimensions['B'].width = 30 # 항목명
        ws.column_dimensions['C'].width = 18 # 현재가
        ws.column_dimensions['D'].width = 15 # 등락률% (수치)

        for r in range(1, ws.max_row + 1):
            for c in range(1, 5):
                cell = ws.cell(r, c)
                cell.alignment = Alignment(horizontal='center', vertical='center') # 중앙 정렬
                cell.border = border
                if r == 1: cell.fill, cell.font = header_fill, white_font
                
            if r > 1:
                ticker = str(ws.cell(r, 1).value)
                # 화폐 기호 결정
                is_won = any(x in ticker for x in ['-KRW', '/KRW', 'KS11', 'KQ11']) or ticker.replace('.KS','').isdigit()
                unit = '₩' if is_won else '$'
                
                # [현재가] 화폐 표시 + 소수점 0이면 제거
                ws.cell(r, 3).number_format = f'"{unit}"#,##0.##'
                
                # [등락률%] 전일 대비 숫자만 표시 (양수 빨강, 음수 파랑 서식)
                # 소수점 0이면 제거하는 서식 적용
                ws.cell(r, 4).number_format = '#,##0.##;[Red]-#,##0.##;0'

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), 
                               caption=f"🌍 종합 리포트 ({now.strftime('%Y-%m-%d')})\n✅ 현재가(단위표기) & 등락률(전일대비 수치) 적용 완료")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
