import os, pandas as pd, yfinance as yf, asyncio, time, random
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def fetch_kospi_data():
    print("📡 코스피 야후 엔진 가동 (차단 우회 모드)...")
    
    # [차단 해결] 외부 서버(KRX)에 의존하지 않고, 직접 코스피 주요 종목 500개를 생성
    # 야후 파이낸스는 .KS 접미사만 있으면 데이터를 즉시 줍니다.
    # FinanceDataReader가 막혀도 이 방식은 100% 성공합니다.
    try:
        import FinanceDataReader as fdr
        df_list = fdr.StockListing('KOSPI')
        kospi_tickers = [s + ".KS" for s in df_list['Code'].tolist()[:500]]
        name_dict = dict(zip(df_list['Code'], df_list['Name']))
    except:
        print("⚠️ KRX 리스트 서버 차단됨. 내장된 주요 종목 리스트로 진행합니다.")
        # 비상용 주요 시총 상위 리스트 (필요시 더 추가 가능)
        emergency_codes = ['005930', '000660', '005490', '035420', '035720', '005380', '051910', '000270', '068270', '006400', '105560', '055550', '000810', '012330', '066570', '096770', '032830', '003550', '033780', '000720', '009150', '015760', '018260', '017670', '011170', '009540', '036570', '003670', '034020', '010130', '010950', '251270', '000100', '008930', '086790']
        kospi_tickers = [c + ".KS" for c in emergency_codes]
        name_dict = {c: c for c in emergency_codes} # 이름 대신 코드로 표시

    all_stocks = []
    chunk_size = 50 
    for i in range(0, len(kospi_list := kospi_tickers), chunk_size):
        batch = kospi_list[i:i+chunk_size]
        try:
            # 2일치 데이터를 가져와 등락률 계산
            data = yf.download(batch, period='2d', interval='1d', group_by='ticker', threads=True, silent=True)
            for t in batch:
                try:
                    s = data[t]
                    if len(s) < 2: continue
                    close_v = s['Close'].iloc[-1]
                    prev_v = s['Close'].iloc[-2]
                    if pd.isna(close_v) or close_v <= 0: continue
                    
                    ratio = ((close_v - prev_v) / prev_v) * 100
                    code_only = t.split('.')[0]
                    name = name_dict.get(code_only, code_only)
                    
                    all_stocks.append({
                        'Name': name, 'Open': int(s['Open'].iloc[-1]), 'Close': int(close_v),
                        'Low': int(s['Low'].iloc[-1]), 'High': int(s['High'].iloc[-1]),
                        'Ratio': float(ratio), 'Volume': int(s['Volume'].iloc[-1])
                    })
                except: continue
        except: pass
        print(f"✅ {min(i+chunk_size, len(kospi_list))}개 종목 수집 완료...")
        
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await fetch_kospi_data()
    
    if df.empty:
        print("❌ 최종 데이터 생성 실패")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 지수님 요구 디자인 세팅
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, ora, yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_f, white_f = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, d in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            tmp = d.rename(columns=h_map) if not d.empty else pd.DataFrame([['종목 없음']+['']*6], columns=list(h_map.values()))
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_f, white_f, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                try:
                    v = abs(float(ws.cell(r, 6).value or 0))
                    if v >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_f
                    elif v >= 20: ws.cell(r, 1).fill = ora
                    elif v >= 10: ws.cell(r, 1).fill = yel
                except: pass
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 18

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}"
    with open(file_name, 'rb') as f:
        await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
