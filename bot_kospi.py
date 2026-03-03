import os, pandas as pd, yfinance as yf, asyncio, time, random
import FinanceDataReader as fdr
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def fetch_stock_data(market_type):
    market_name = "KOSPI" if market_type == 0 else "KOSDAQ"
    market_code = "KOSPI" if market_type == 0 else "KOSDAQ"
    suffix = ".KS" if market_type == 0 else ".KQ"
    
    print(f"📡 {market_name} 야후 엔진 전수조사 시작 (FDR 리스트 모드)...")
    
    try:
        # KRX 대신 FinanceDataReader로 리스트 확보 (차단 없음)
        df_list = fdr.StockListing(market_code)
        # 상위 600개 종목 추출 (거래소 안정성을 위해)
        stock_list = df_list['Code'].tolist()[:600]
        tickers = [s + suffix for s in stock_list]
        # 종목명 매핑용 딕셔너리
        name_dict = dict(zip(df_list['Code'], df_list['Name']))
    except Exception as e:
        print(f"❌ 종목 리스트 획득 실패: {e}")
        return pd.DataFrame()

    all_stocks = []
    chunk_size = 50 
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            data = yf.download(batch, period='2d', interval='1d', group_by='ticker', threads=True, silent=True)
            for t in batch:
                try:
                    s = data[t]
                    if len(s) < 2: continue
                    close_v = s['Close'].iloc[-1]
                    prev_v = s['Close'].iloc[-2]
                    if close_v <= 0 or pd.isna(close_v): continue
                    
                    ratio = ((close_v - prev_v) / prev_v) * 100
                    code_only = t.split('.')[0]
                    name = name_dict.get(code_only, t)
                    
                    all_stocks.append({
                        'Name': name, 'Open': int(s['Open'].iloc[-1]), 'Close': int(close_v),
                        'Low': int(s['Low'].iloc[-1]), 'High': int(s['High'].iloc[-1]),
                        'Ratio': float(ratio), 'Volume': int(s['Volume'].iloc[-1])
                    })
                except: continue
        except: pass
        print(f"✅ {min(i+chunk_size, len(tickers))}개 분석 완료...")
        
    return pd.DataFrame(all_stocks)

# 이하 main() 함수 및 디자인 로직은 기존과 동일 (생략 가능하나 확인용으로 포함)
async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await fetch_stock_data(0) # 0: KOSPI
    if df.empty: return
    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)
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
    with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__": asyncio.run(main())
