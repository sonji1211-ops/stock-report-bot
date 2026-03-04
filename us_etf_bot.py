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
        
        ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
        diff_value = last_c - prev_c
            
        return {
            '티커': symbol, 
            '항목명': ASSET_NAMES.get(symbol, symbol), 
            '현재가': last_c, 
            '전일대비': diff_value,
            '등락률': ratio
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
        df[['티커', '항목명', '현재가', '전일대비']].to_excel(writer, sheet_name='현황', index=False)
        ws = writer.sheets['현황']
        
        h_fill = PatternFill(start_color="444444", end_color="444444", fill_type="solid")
        f_white = Font(color="FFFFFF", bold=True)
        # ETF 최적화 색상
        colors = {
            'red': PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),   # 10%↑
            'orange': PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid"),# 5%↑
            'yellow': PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid") # 3%↑
        }
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 15

        for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=4), 1):
            for c_idx, cell in enumerate(row, 1):
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                if r_idx == 1:
                    cell.fill, cell.font = h_fill, f_white
                
            if r_idx > 1:
                ratio_val = df.iloc[r_idx-2]['등락률']
                abs_ratio = abs(ratio_val)
                name_cell = ws.cell(row=r_idx, column=2)
                
                # [ETF 전용 기준 변경] 3% / 5% / 10%
                if abs_ratio >= 10: 
                    name_cell.fill, name_cell.font = colors['red'], f_white
                elif abs_ratio >= 5: 
                    name_cell.fill = colors['orange']
                elif abs_ratio >= 3: 
                    name_cell.fill = colors['yellow']

                ticker = str(ws.cell(r_idx, 1).value)
                is_won = any(x in ticker for x in ['-KRW', '/KRW', 'KS11', 'KQ11']) or ticker.replace('.KS','').isdigit()
                unit = '₩' if is_won else '$'
                ws.cell(r_idx, 3).number_format = f'"{unit}"#,##0.##'
                ws.cell(r_idx, 4).number_format = '#,##0.##;[Red]-#,##0.##;0'

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), 
                               caption=f"🌍 종합 리포트 ({now.strftime('%Y-%m-%d')})\n✅ ETF 맞춤형 강조 (3%🟡 5%🟠 10%🔴) 적용")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
