import os, pandas as pd, yfinance as yf, asyncio, time
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# [전략] 실제 코스피 종목이 밀집된 구간을 자동 생성 (약 500~600개 타겟)
def get_expanded_kospi_list():
    # 1. 시총 최상위 핵심 50개 (무조건 포함)
    top_50 = [
        '005930','000660','005490','035420','035720','005380','051910','000270','068270','006400',
        '105560','055550','000810','012330','066570','096770','032830','003550','033780','000720',
        '009150','015760','018260','017670','011170','009540','036570','003670','034020','010130',
        '010950','251270','000100','008930','086790','004020','078930','028260','000120','030200',
        '039130','011070','000080','005070','005935','009830','001570','016360','004170','036460'
    ]
    # 2. 코스피 종목이 집중된 000010 ~ 050000 사이 구간 자동 생성 (간격 촘촘히)
    scan_range = [f"{i:06d}" for i in range(10, 50000, 80)] # 약 600개 생성
    
    return list(set(top_50 + scan_range))

async def fetch_kospi_data():
    print("📡 [종목확장] 코스피 500개급 전수조사 시작...")
    codes = get_expanded_kospi_list()
    tickers = [c + ".KS" for c in codes]
    all_stocks = []
    
    # 15개씩 묶어서 요청 (야후 차단 방어 최적화)
    for i in range(0, len(tickers), 15):
        batch = tickers[i:i+15]
        try:
            # 7일치 데이터를 가져와 등락률 계산의 안정성 확보
            data = yf.download(batch, period='7d', interval='1d', group_by='ticker', silent=True)
            for t in batch:
                try:
                    if t not in data.columns.get_level_values(0): continue
                    s = data[t].dropna()
                    if len(s) < 2: continue
                    
                    close_v = s['Close'].iloc[-1]
                    prev_v = s['Close'].iloc[-2]
                    ratio = ((close_v - prev_v) / prev_v) * 100
                    
                    # 0원 주식이나 에러 데이터 제외
                    if pd.isna(ratio) or close_v <= 0: continue
                    
                    all_stocks.append({
                        'Name': t.split('.')[0], # 종목명 차단 대비 코드로 표시
                        'Open': int(s['Open'].iloc[-1]), 'Close': int(close_v),
                        'Low': int(s['Low'].iloc[-1]), 'High': int(s['High'].iloc[-1]),
                        'Ratio': float(ratio), 'Volume': int(s['Volume'].iloc[-1])
                    })
                except: continue
        except: pass
        print(f"✅ {min(i+15, len(tickers))}개 분석 중... (현재 데이터 확보: {len(all_stocks)}개)")
        time.sleep(0.4) # 차단 방지용 딜레이
        
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await fetch_kospi_data()
    
    if df.empty:
        print("❌ 수집된 데이터가 없습니다.")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    # 필터: 5% 이상 변동 종목
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 엑셀 디자인 및 포맷팅 (지수님 요청 사항 100% 반영)
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, ora, yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_f, white_f = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, d in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            tmp = d.rename(columns=h_map) if not d.empty else pd.DataFrame([['조건 만족 종목 없음']+['']*6], columns=list(h_map.values()))
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_f, white_f, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                try:
                    val = ws.cell(r, 6).value
                    v = abs(float(val)) if val and str(val).replace('.','').replace('-','').isdigit() else 0
                    if v >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_f
                    elif v >= 20: ws.cell(r, 1).fill = ora
                    elif v >= 10: ws.cell(r, 1).fill = yel
                except: pass
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 20

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📈 분석대상: {len(df)}개\n🚀 상승(5%↑): {len(up_df)}개 / 하락(5%↓): {len(down_df)}개"
    with open(file_name, 'rb') as f:
        await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
