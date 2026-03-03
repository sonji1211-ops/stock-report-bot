import os, pandas as pd, requests, re, io, time, random, asyncio
from datetime import datetime, timedelta
from bs4 import BeautifulSoup
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def fetch_naver_page(url):
    headers = {
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/sise/sise_market_sum.naver?sosok=1',
        'Cookie': 'NID_AUT=dummy; NID_SES=dummy;'
    }
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        return resp.text if resp.status_code == 200 else None
    except: return None

async def get_kosdaq_scan():
    fields = "field=quant&field=open&field=high&field=low&field=frate"
    all_stocks = []
    init_html = fetch_naver_page("https://finance.naver.com/sise/sise_market_sum.naver?sosok=1")
    last_page = int(max(map(int, re.findall(r'page=(\d+)', init_html)))) if init_html else 1
    
    print(f"📡 KOSDAQ 정밀 전수조사 시작 ({last_page}페이지)...")

    for page in range(1, last_page + 1):
        # 코스닥은 종목이 더 많으므로 10페이지마다 5초 휴식
        if page % 10 == 0:
            print("☕ 네이버 눈피하기... 5초간 휴식")
            time.sleep(5)
            
        url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok=1&{fields}&page={page}"
        html = fetch_naver_page(url)
        
        if html and "종목명" in html:
            soup = BeautifulSoup(html, 'html.parser')
            table = soup.find('table', {'class': 'type_2'})
            if table:
                df = pd.read_html(io.StringIO(str(table)))[0].dropna(subset=['종목명'])
                for col in ['등락률', '현재가', '시가', '고가', '저가', '거래량']:
                    df[col] = pd.to_numeric(df[col].astype(str).str.replace(r'[%,\+]', '', regex=True), errors='coerce').fillna(0)
                for _, row in df.iterrows():
                    all_stocks.append({'Name': row['종목명'], 'Open': int(row['시가']), 'Close': int(row['현재가']), 
                                       'Low': int(row['저가']), 'High': int(row['고가']), 'Ratio': float(row['등락률']), 'Volume': int(row['거래량'])})
            time.sleep(random.uniform(1.2, 2.5)) # 코스닥은 더 천천히
        else:
            print(f"🛑 {page}p 차단 발생! 현재까지 데이터만 전송합니다.")
            break 
            
        if page % 10 == 0: print(f"✅ KOSDAQ {page}p 완료")
            
    return pd.DataFrame(all_stocks)

async def send_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await get_kosdaq_scan()
    if df.empty: return

    report_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSDAQ_{report_type}.xlsx"
    
    up_df = df[(df['Ratio'] >= 5.0) & (df['Volume'] > 0)].sort_values('Ratio', ascending=False)
    down_df = df[(df['Ratio'] <= -5.0) & (df['Volume'] > 0)].sort_values('Ratio', ascending=True)
    
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, orange, yellow = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_fill, white_font = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in {'코스닥_상승': up_df, '코스닥_하락': down_df}.items():
            data.rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_fill, white_font, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                val = abs(float(ws.cell(r, 6).value or 0))
                if val >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_font
                elif val >= 20: ws.cell(r, 1).fill = orange
                elif val >= 10: ws.cell(r, 1).fill = yellow
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 18

    msg = (f"📅 {now.strftime('%m-%d')} *[KOSDAQ {report_type}]*\n"
           f"📊 분석: 코스닥 전수조사 ({len(df)}개)\n"
           f"📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}\n"
           f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑ (하락주 포함)")
    
    with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    os.remove(file_name)

if __name__ == "__main__": asyncio.run(send_report())
