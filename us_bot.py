import os, pandas as pd, asyncio, time, datetime
import yfinance as yf
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정] 
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_verified_tickers():
    """야후에서 확실히 응답하는 실존 우량주/활성주 리스트 (추측성 난수 0%)"""
    # KOSPI 200 주요 종목
    kospi = ["005930", "000660", "005380", "035420", "035720", "005490", "051910", "006400", "000270", "068270",
             "012330", "010130", "033780", "009150", "034730", "018260", "000810", "015760", "032830", "003550",
             "000100", "000120", "000720", "000880", "001040", "011780", "011790", "012450", "017670", "020150"]
    # KOSDAQ 150 및 주요 종목
    kosdaq = ["247540", "086520", "091990", "066970", "293490", "028300", "058470", "214150", "035900", "036830",
              "112040", "131390", "145020", "196170", "204320", "214370", "230360", "253450", "263750", "272210"]
    
    # 지수님, 여기에 실제 상장된 다른 번호들을 추가하면 분석 범위가 늘어납니다.
    # 지금은 '안전'이 제일이므로 확실한 놈들 위주로 구성했습니다.
    return [f"{c}.KS" for c in kospi] + [f"{c}.KQ" for c in kosdaq]

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    tickers = get_verified_tickers()
    print(f"📡 [안전 모드] {len(tickers)}개 확정 종목 분석 시작...")

    collected = []
    # 차단 방지를 위해 10개씩 조심스럽게 호출
    for i in range(0, len(tickers), 10):
        batch = tickers[i:i+10]
        try:
            # 2일치 데이터를 가져와 변동성 계산 (주말/공휴일 대응)
            data = yf.download(batch, period="2d", interval="1d", group_by='ticker', progress=False)
            
            for t in batch:
                if t not in data.columns.levels[0]: continue
                df_t = data[t].dropna()
                if len(df_t) < 2: continue
                
                close_today = df_t['Close'].iloc[-1]
                close_prev = df_t['Close'].iloc[-2]
                ratio = ((close_today - close_prev) / close_prev) * 100
                volume = df_t['Volume'].iloc[-1]
                
                if volume > 0: # 실제 거래된 종목만
                    collected.append({
                        'Code': t.split('.')[0], 'Name': t.split('.')[0], # 야후는 사명을 안주므로 일단 코드로
                        'Market': "KOSPI" if t.endswith(".KS") else "KOSDAQ",
                        'Open': int(df_t['Open'].iloc[-1]), 'Close': int(close_today),
                        'Low': int(df_t['Low'].iloc[-1]), 'High': int(df_t['High'].iloc[-1]),
                        'Ratio': float(ratio), 'Volume': int(volume)
                    })
        except: continue
        time.sleep(0.5) # 정중한 대기
        print(f"📦 {min(i+10, len(tickers))}개 완료...")

    if not collected: return

    df_final = pd.DataFrame(collected)
    # 지수님 요청: ±5% 필터
    final_filtered = df_final[(df_final['Ratio'] >= 5) | (df_final['Ratio'] <= -5)]

    # 엑셀 시트 분류 및 디자인
    file_name = f"{now.strftime('%m%d')}_안정형_리포트.xlsx"
    h_fill, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    p_red, p_ora, p_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                sub = final_filtered[(final_filtered['Market']==m) & ((final_filtered['Ratio']>0) if trend=='상승' else (final_filtered['Ratio']<0))]
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market'])
                
                s_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]

                # 디자인 입히기
                for cell in ws[1]: # 헤더
                    cell.fill, cell.font, cell.alignment, cell.border = h_fill, f_white, Alignment(horizontal='center'), border
                for r in range(2, ws.max_row + 1): # 본문
                    rv = abs(float(ws.cell(r, 7).value or 0))
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment, cell.border = Alignment(horizontal='center'), border
                        if c in [3,4,5,6,8]: cell.number_format = '#,##0' # 숫자 콤마
                        if c == 7: cell.number_format = '0.00' # 등락률
                        if c == 2: # 강조 색상
                            if rv >= 28: cell.fill, cell.font = p_red, Font(color="FFFFFF", bold=True)
                            elif rv >= 20: cell.fill = p_ora
                            elif rv >= 10: cell.fill = p_yel
                ws.column_dimensions['B'].width = 15

    # 전송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 야후기반 리포트\n\n"
               f"📊 분석: {len(df_final)}개 / 필터통과: {len(final_filtered)}개\n"
               f"💡 네이버 차단 우회 및 확정 리스트 분석 완료")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(send_smart_report())
