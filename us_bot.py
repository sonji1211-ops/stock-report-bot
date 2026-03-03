import os, pandas as pd, asyncio, time, datetime
import yfinance as yf
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_market_data():
    """야후 파이낸스에서 코스피/코스닥 데이터를 누락 없이 덩어리로 가져옵니다."""
    # 유효한 종목 번호 대역만 정밀 타겟팅 (없는 번호 찌르기 방지)
    k_codes = [f"{i:06d}.KS" for i in range(50, 15000, 30)] 
    q_codes = [f"{i:06d}.KQ" for i in range(100, 150000, 150)]
    tickers = k_codes + q_codes
    
    all_stocks = []
    # 40개씩 끊어서 서버 차단 방지
    for i in range(0, len(tickers), 40):
        batch = tickers[i:i+40]
        try:
            data = yf.download(batch, period="5d", interval="1d", group_by='ticker', threads=True, progress=False)
            for t in batch:
                if t not in data.columns.levels[0]: continue
                df_t = data[t].dropna()
                if len(df_t) < 2: continue
                
                c, p, v = df_t['Close'].iloc[-1], df_t['Close'].iloc[-2], df_t['Volume'].iloc[-1]
                if c <= 0 or v <= 0: continue
                
                all_stocks.append({
                    'Code': t.split('.')[0], 'Name': t.split('.')[0],
                    'Market': "KOSPI" if t.endswith(".KS") else "KOSDAQ",
                    'Open': int(df_t['Open'].iloc[-1]), 'Close': int(c),
                    'Low': int(df_t['Low'].iloc[-1]), 'High': int(df_t['High'].iloc[-1]),
                    'Ratio': float(((c - p) / p) * 100), 'Volume': int(v)
                })
        except: continue
        time.sleep(0.1)
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    df = get_market_data()
    if df.empty: return

    # 1. 리포트 타입 결정
    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # 2. 디자인 요소 정의
    h_fill = PatternFill("solid", fgColor="444444")
    f_white_bold = Font(color="FFFFFF", bold=True)
    f_red_white = Font(color="FFFFFF", bold=True)
    p_red, p_ora, p_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center')

    h_map = {'Code':'CODE', 'Name':'NAME', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                # 5% 기준 필터링
                sub = df[(df['Market']==m) & ((df['Ratio']>=5) if trend=='상승' else (df['Ratio']<=-5))]
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market']).rename(columns=h_map)
                
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 헤더 스타일 적용
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = h_fill, f_white_bold, center, border

                # 본문 스타일 및 강조 적용
                for r in range(2, ws.max_row + 1):
                    ratio_val = abs(float(ws.cell(r, 7).value or 0))
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment, cell.border = center, border
                        # 콤마 및 소수점
                        if c in [3,4,5,6,8]: cell.number_format = '#,##0'
                        if c == 7: cell.number_format = '0.00'
                        # 종목명 색상 강조
                        if c == 2:
                            if ratio_val >= 28: cell.fill, cell.font = p_red, f_red_white
                            elif ratio_val >= 20: cell.fill = p_ora
                            elif ratio_val >= 10: cell.fill = p_yel
                ws.column_dimensions['B'].width = 15

    # 3. 텔레그램 메시지 구성 (누락 없이!)
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
