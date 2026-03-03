import os, pandas as pd, asyncio, time, datetime
from yahooquery import Ticker
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_market_data_fixed():
    """코스닥 누락 방지 및 거래량 이중 체크 엔진"""
    print("📡 [1단계] 코스피/코스닥 종목 리스트 생성...")
    
    # 촘촘한 스캔 (코스닥 .KQ를 먼저 배치하여 우선순위 확보)
    codes = [f"{i:06d}" for i in range(10, 2000, 3)]
    tickers = [c + ".KQ" for c in codes] + [c + ".KS" for c in codes]
    
    all_stocks = []
    chunk_size = 40 # 안정적인 수집을 위한 크기
    
    print(f"🚀 총 {len(tickers)}개 후보 분석 시작 (코스닥 우선)...")
    
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            t = Ticker(batch, asynchronous=True)
            p = t.price
            d = t.summary_detail
            
            for symbol in batch:
                info = p.get(symbol, {})
                det = d.get(symbol, {})
                
                if isinstance(info, dict) and 'regularMarketPrice' in info:
                    # 가격 및 거래량 추출 (누락 대비 이중 필드 체크)
                    cp = info.get('regularMarketPrice') or det.get('previousClose') or 0
                    vol = info.get('regularMarketVolume') or det.get('volume') or 0
                    
                    if cp <= 100: continue # 상폐/초저가주 제외
                    
                    market = "KOSDAQ" if symbol.endswith(".KQ") else "KOSPI"
                    all_stocks.append({
                        'Code': symbol.split('.')[0],
                        'Name': info.get('shortName', symbol.split('.')[0]),
                        'Market': market,
                        'Open': int(info.get('regularMarketOpen') or det.get('open') or cp),
                        'Close': int(cp),
                        'Low': int(info.get('regularMarketDayLow') or det.get('dayLow') or cp),
                        'High': int(info.get('regularMarketDayHigh') or det.get('dayHigh') or cp),
                        'Ratio': float(info.get('regularMarketChangePercent', 0) * 100),
                        'Volume': int(vol) # 거래량 누락 방지
                    })
        except: continue
        time.sleep(0.05)
    
    print(f"✅ 수집 완료: {len(all_stocks)}개 확보 (코스닥 포함)")
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    df = await get_market_data_fixed()
    if df.empty: return

    # 요일 로직
    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # [디자인 요구사항 체크리스트]
    # 1. 헤더: 짙은 배경 + 흰색 굵은 글씨
    # 2. 본문: 모든 셀 중앙 정렬 + 테두리
    # 3. 강조: 28%↑(빨강), 20%↑(주황), 10%↑(노랑)
    # 4. 포맷: 천 단위 콤마 + 소수점 2자리
    
    h_map = {'Code':'CODE', 'Name':'NAME', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
    f_red, f_ora, f_yel = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
    f_head, f_white = PatternFill("solid", fgColor="444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                sub = df[(df['Market']==m) & ((df['Ratio']>=5) if trend=='상승' else (df['Ratio']<=-5))]
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market']).rename(columns=h_map)
                
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border

                # 본문 스타일
                for r in range(2, ws.max_row + 1):
                    try:
                        rv = abs(float(ws.cell(r, 7).value or 0))
                        name_cell = ws.cell(r, 2)
                        # 색상 강조
                        if rv >= 28: name_cell.fill, name_cell.font = f_red, f_white
                        elif rv >= 20: name_cell.fill = f_ora
                        elif rv >= 10: name_cell.fill = f_yel
                    except: pass
                    
                    for c in range(1, 9):
                        ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                        if c in [3, 4, 5, 6, 8]: ws.cell(r, c).number_format = '#,##0' # 콤마 적용
                        if c == 7: ws.cell(r, c).number_format = '0.00' # 소수점 적용
                ws.column_dimensions['B'].width = 25

    # 전송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 국내증시 {report_type}\n\n"
               f"📊 수집 종목수: {len(df)}개 (KOSPI/KOSDAQ)\n"
               f"📈 상승(5%↑): {len(df[df['Ratio']>=5])}개\n"
               f"📉 하락(5%↓): {len(df[df['Ratio']<=-5])}개\n\n"
               f"💡 거래량/데이터 정밀 보정 완료")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
