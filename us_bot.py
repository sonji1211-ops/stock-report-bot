import os, pandas as pd, asyncio, time, datetime
import yfinance as yf
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_final_korea_data():
    print("📡 [1단계] 야후 파이낸스 직접 쿼리 시작 (KRX/네이버 우회)...")
    
    # 아까 에러 난 000... 대역 대신, 실제 우량주가 몰려있는 대역으로 정밀 타겟팅
    # 삼성전자(005930), 현대차(005380) 등 005~009 대역과 코스닥 주요 대역
    target_ranges = [
        [f"{i:06d}.KS" for i in range(5000, 9999, 5)],   # 코스피 주요 대역
        [f"{i:06d}.KQ" for i in range(30000, 150000, 50)] # 코스닥 주요 대역
    ]
    tickers = [item for sublist in target_ranges for item in sublist]
    
    print(f"🚀 총 {len(tickers)}개 후보 분석 시작... (차단 회피 모드)")
    
    try:
        # threads=True로 속도를 높이고, 데이터를 2일치만 가져와서 부하를 줄입니다.
        raw = yf.download(tickers, period="2d", interval="1d", group_by='ticker', threads=True)
    except:
        return pd.DataFrame()

    all_stocks = []
    for ticker in tickers:
        try:
            if ticker not in raw.columns.levels[0]: continue
            df_t = raw[ticker].dropna()
            if len(df_t) < 2: continue
            
            curr_c = df_t['Close'].iloc[-1]
            prev_c = df_t['Close'].iloc[-2]
            vol = df_t['Volume'].iloc[-1]
            
            if curr_c <= 100 or vol < 1000: continue # 동전주 및 거래량 없는 주식 제외
            
            ratio = ((curr_c - prev_c) / prev_c) * 100
            market = "KOSPI" if ticker.endswith(".KS") else "KOSDAQ"
            
            # 야후에서 제공하는 영문명이라도 가져옵니다 (차단된 국문명 대신)
            all_stocks.append({
                'Code': ticker.split('.')[0],
                'Name': ticker.split('.')[0], # 일단 코드로 넣고, 엑셀에서 확인
                'Market': market,
                'Open': int(df_t['Open'].iloc[-1]),
                'Close': int(curr_c),
                'Low': int(df_t['Low'].iloc[-1]),
                'High': int(df_t['High'].iloc[-1]),
                'Ratio': float(ratio),
                'Volume': int(vol)
            })
        except: continue

    print(f"✅ 수집 완료: {len(all_stocks)}개 유효 종목 확보")
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    df = get_final_korea_data()
    
    if df.empty:
        print("🚨 수집된 데이터가 없습니다.")
        return

    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # [지수님 디자인 툴킷 적용]
    h_map = {'Code':'CODE', 'Name':'NAME', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
    f_red, f_ora, f_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    f_head, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                # 5% 기준 필터링 (없으면 3%로 하향)
                sub = df[(df['Market']==m) & ((df['Ratio']>=5) if trend=='상승' else (df['Ratio']<=-5))]
                if len(sub) < 3:
                    sub = df[(df['Market']==m) & ((df['Ratio']>=3) if trend=='상승' else (df['Ratio']<=-3))]
                
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market']).rename(columns=h_map)
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 헤더 스타일 (#444444)
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border

                # 본문 스타일 (중앙정렬, 콤마, 테두리, 강조색)
                for r in range(2, ws.max_row + 1):
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
                ws.column_dimensions['B'].width = 15

    # 텔레그램 전송
    async with bot:
        msg = f"📅 {now.strftime('%Y-%m-%d')} {report_type} 리포트\n📊 분석 종목: {len(df)}개\n📈 상승(급등): {len(df[df['Ratio']>=5])}개"
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
