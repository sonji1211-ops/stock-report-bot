import os, pandas as pd, requests, re, io, time, random, asyncio
from datetime import datetime, timedelta
from telegram import Bot
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def fetch_naver_stock(sosok, page):
    url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&field=quant&field=open&field=high&field=low&field=frate&page={page}"
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0 Safari/537.36', 'Referer': 'https://finance.naver.com/'}
    try:
        resp = requests.get(url, headers=headers, timeout=10)
        if resp.status_code != 200: return []
        soup = BeautifulSoup(resp.text, 'lxml')
        table = soup.find('table', {'class': 'type_2'})
        if not table: return []
        
        rows = []
        for tr in table.find_all('tr'):
            tds = tr.find_all('td')
            if len(tds) < 10: continue
            name = tds[1].get_text(strip=True)
            if not name: continue
            def clean(i): return tds[i].get_text(strip=True).replace(',', '').replace('%', '').replace('+', '')
            try:
                rows.append({'Name': name, 'Close': int(clean(2)), 'Ratio': float(clean(4)), 'Volume': int(clean(5)),
                             'Open': int(clean(7)), 'High': int(clean(8)), 'Low': int(clean(9))})
            except: continue
        return rows
    except: return []

async def run_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    all_data = []
    
    print("📡 KOSPI 전수조사 시작 (최대 30p)...")
    for p in range(1, 31):
        data = fetch_naver_stock(0, p) # sosok=0 (KOSPI)
        if not data: break
        all_data.extend(data)
        if p % 10 == 0: print(f"✅ {p}p 완료 ({len(all_data)}개)")
        time.sleep(random.uniform(0.3, 0.6))

    df = pd.DataFrame(all_data)
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
            tmp = d.rename(columns=h_map) if not d.empty else pd.DataFrame([['조건 만족 종목 없음']+['']*6], columns=list(h_map.values()))
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

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📊 수집: {len(df)}개\n📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}"
    with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__": asyncio.run(run_report())
