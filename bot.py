import os, pandas as pd, asyncio, time
from yahooquery import Ticker
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_high_volume_data():
    """종목수 극대화 엔진: 유효한 종목 대역을 촘촘하게 스캔"""
    print("📡 [1단계] 종목 스캔 엔진 가동 (KOSPI/KOSDAQ)...")
    
    # 촘촘한 스캔을 위해 주요 대역 설정 (간격 좁힘)
    scan_ranges = [
        (1, 1500, 3),    # 주요 종목 대역
        (2000, 3500, 3),  # 중소형주 대역
        (5000, 9000, 5)   # 기타 대역
    ]
    
    tickers = []
    for start, end, step in scan_ranges:
        for i in range(start, end, step):
            code = f"{i:06d}"
            tickers.append(code + ".KS") # 코스피
            tickers.append(code + ".KQ") # 코스닥

    all_stocks = []
    chunk_size = 40 # 야후 쿼리 최적화 크기
    
    print(f"🚀 총 {len(tickers)}개 후보 종목 분석 시작...")
    
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            t = Ticker(batch, asynchronous=True)
            p_data = t.price
            d_data = t.summary_detail
            
            for symbol in batch:
                p = p_data.get(symbol, {})
                d = d_data.get(symbol, {})
                
                # 데이터 유효성 검사 (가격이 있어야 살아있는 종목)
                if isinstance(p, dict) and 'regularMarketPrice' in p:
                    cp = p.get('regularMarketPrice') or d.get('previousClose') or 0
                    if cp <= 100: continue # 초저가주(동전주 일부) 제외로 노이즈 제거
                    
                    market = "KOSPI" if symbol.endswith(".KS") else "KOSDAQ"
                    all_stocks.append({
                        'Code': symbol.split('.')[0],
                        'Name': p.get('shortName', symbol.split('.')[0]),
                        'Market': market,
                        'Open': int(p.get('regularMarketOpen') or d.get('open') or cp),
                        'Close': int(cp),
                        'Low': int(p.get('regularMarketDayLow') or d.get('dayLow') or cp),
                        'High': int(p.get('regularMarketDayHigh') or d.get('dayHigh') or cp),
                        'Ratio': float(p.get('regularMarketChangePercent', 0) * 100),
                        'Volume': int(p.get('regularMarketVolume') or d.get('volume') or 0)
                    })
        except: continue
        # if i % 400 == 0: print(f" 진행 중... ({i}/{len(tickers)})")
    
    print(f"✅ 수집 완료: {len(all_stocks)}개 유효 종목 확보")
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    
    df = await get_high_volume_data()
    if df.empty: return

    # 요일 로직
    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # 디자인 설정
    h_map = {'Code':'CODE', 'Name':'NAME', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
    f_red, f_ora, f_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    f_head, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                # 필터링 (등락률 5% 이상/이하)
                sub = df[(df['Market']==m) & ((df['Ratio']>=5) if trend=='상승' else (df['Ratio']<=-5))]
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market']).rename(columns=h_map)
                
                s_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]

                # 엑셀 스타일링
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border

                for r in range(2, ws.max_row + 1):
                    # 색상 강조 로직
                    try:
                        rv = abs(float(ws.cell(r, 7).value or 0))
                        name_cell = ws.cell(r, 2)
                        if rv >= 28: name_cell.fill, name_cell.font = f_red, f_white
                        elif rv >= 20: name_cell.fill = f_ora
                        elif rv >= 10: name_cell.fill = f_yel
                    except: pass
                    
                    for c in range(1, 9):
                        ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                        if c in [3, 4, 5, 6, 8]: ws.cell(r, c).number_format = '#,##0'
                        if c == 7: ws.cell(r, c).number_format = '0.00'
                ws.column_dimensions['B'].width = 25

    # 텔레그램 전송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 국내증시 {report_type} 리포트\n\n"
               f"📊 종목수: {len(df)}개 스캔 완료\n"
               f"📈 상승(5%↑): {len(df[df['Ratio']>=5])}개\n"
               f"📉 하락(5%↓): {len(df[df['Ratio']<=-5])}개\n\n"
               f"💡 🔴28%↑ 🟠20%↑ 🟡10%↑")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
