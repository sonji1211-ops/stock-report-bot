import os, pandas as pd, requests, re, io, time, random, asyncio
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def fetch_naver_page(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/'
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        return resp.text if resp.status_code == 200 else None
    except: return None

async def get_kospi_scan():
    fields = "field=quant&field=open&field=high&field=low&field=frate"
    all_stocks = []
    init_html = fetch_naver_page("https://finance.naver.com/sise/sise_market_sum.naver?sosok=0")
    if not init_html: return pd.DataFrame()
    last_page = int(max(map(int, re.findall(r'page=(\d+)', init_html)))) if re.findall(r'page=(\d+)', init_html) else 1
    target_page = min(last_page, 30)
    
    print(f"📡 KOSPI 집중 전수조사 시작 (상위 {target_page}p)...")
    for page in range(1, target_page + 1):
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok=0&{fields}&page={page}"
        html = fetch_naver_page(url)
        if html and "종목명" in html:
            try:
                dfs = pd.read_html(io.StringIO(html))
                # [보강] 종목명이 포함된 진짜 주식 표 찾기
                df = next((d for d in dfs if '종목명' in d.columns), None)
                if df is not None:
                    df = df.dropna(subset=['종목명']).copy()
                    df.columns = [str(c).strip() for c in df.columns]
                    for col in ['등락률', '현재가', '시가', '고가', '저가', '거래량']:
                        if col in df.columns:
                            df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[%,\+]', '', regex=True), errors='coerce').fillna(0)
                    for _, row in df.iterrows():
                        if row['현재가'] > 0:
                            all_stocks.append({'Name':str(row['종목명']), 'Open':int(row['시가']), 'Close':int(row['현재가']), 
                                               'Low':int(row['저가']), 'High':int(row['고가']), 'Ratio':float(row['등락률']), 'Volume':int(row['거래량'])})
                time.sleep(random.uniform(0.5, 1.0))
            except: continue
        else: break
        if page % 10 == 0: print(f"✅ KOSPI {page}p 완료")
    return pd.DataFrame(all_stocks)

async def send_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await get_kospi_scan()
    if df.empty: return

    report_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{report_type}.xlsx"
    
    up_df = df[(df['Ratio'] >= 5.0) & (df['Volume'] > 0)].sort_values('Ratio', ascending=False)
    down_df = df[(df['Ratio'] <= -5.0) & (df['Volume'] > 0)].sort_values('Ratio', ascending=True)
    
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, orange, yellow = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_fill, white_font = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            if data.empty:
                data = pd.DataFrame([['해당 종목 없음'] + [''] * 6], columns=list(h_map.values()))
            else:
                data = data.rename(columns=h_map)
                
            data.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_fill, white_font, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                try:
                    val = abs(float(ws.cell(r, 6).value or 0))
                    if val >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_font
                    elif val >= 20: ws.cell(r, 1).fill = orange
                    elif val >= 10: ws.cell(r, 1).fill = yellow
                except: pass
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 18

    msg = (f"📅 {now.strftime('%m-%d')} *[KOSPI {report_type}]*\n"
           f"📊 분석: 상위 30p ({len(df)}개)\n"
           f"📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}\n"
           f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
    
    try:
        with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    except Exception as e: print(f"❌ 전송 실패: {e}")
    finally:
        if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__": asyncio.run(send_report())
