import os, pandas as pd, requests, re, io, time, random, asyncio
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def fetch_page(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0 Safari/537.36', 'Referer': 'https://finance.naver.com/'}
    try:
        resp = requests.get(url, headers=headers, timeout=7) # 타임아웃 단축
        return resp.text if resp.status_code == 200 else None
    except: return None

async def get_data():
    all_stocks = []
    # KOSPI 종목 리스트 (sosok=0)
    init_html = fetch_page("https://finance.naver.com/sise/sise_market_sum.naver?sosok=0")
    if not init_html: return pd.DataFrame()
    last_p = min(int(max(re.findall(r'page=(\d+)', init_html), default=1)), 30)
    
    print(f"📡 KOSPI 수집 시작 (목표: {last_p}p)...")
    for p in range(1, last_p + 1):
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok=0&field=quant&field=open&field=high&field=low&field=frate&page={p}"
        html = fetch_page(url)
        if not html or "종목명" not in html: 
            print(f"⚠️ {p}p에서 차단 감지. 수집 종료 및 결과 전송."); break
        try:
            dfs = pd.read_html(io.StringIO(html))
            df = next((d for d in dfs if '종목명' in d.columns), None)
            if df is not None:
                df = df.dropna(subset=['종목명']).copy()
                df.columns = [str(c).strip() for c in df.columns]
                for c_name in ['등락률', '현재가', '시가', '고가', '저가', '거래량']:
                    if c_name in df.columns:
                        df[c_name] = pd.to_numeric(df[c_name].astype(str).str.replace(r'[%,\+]', '', regex=True), errors='coerce').fillna(0)
                for _, row in df.iterrows():
                    if row['현재가'] > 0:
                        all_stocks.append({'Name':str(row['종목명']), 'Open':int(row['시가']), 'Close':int(row['현재가']), 'Low':int(row['저가']), 'High':int(row['고가']), 'Ratio':float(row['등락률']), 'Volume':int(row['거래량'])})
            time.sleep(random.uniform(0.3, 0.7)) # 속도 최적화
        except: break
        if p % 10 == 0: print(f"✅ {p}p 완료 ({len(all_stocks)}종목)")
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await get_data()
    if df.empty: print("❌ 데이터 없음"); return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)
    
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, ora, yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_f, white_font = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            tmp = data.rename(columns=h_map) if not data.empty else pd.DataFrame([['종목 없음']+['']*6], columns=list(h_map.values()))
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_f, white_font, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                try:
                    v = abs(float(ws.cell(r, 6).value or 0))
                    if v >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_font
                    elif v >= 20: ws.cell(r, 1).fill = ora
                    elif v >= 10: ws.cell(r, 1).fill = yel
                except: pass
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 18

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📊 수집: {len(df)}개\n📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}\n💡 🔴28% 🟠20% 🟡10%"
    with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__": asyncio.run(main())
