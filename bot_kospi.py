import os, pandas as pd, requests, time, random, asyncio
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def fetch_api_data(sosok):
    all_stocks = []
    # 네이버 내부 API: sosok 0=코스피, 1=코스닥
    # 차단을 피하기 위해 브라우저가 호출하는 실제 API 주소 사용
    base_url = "https://finance.naver.com/sise/sise_market_sum.naver"
    
    print(f"📡 {'KOSPI' if sosok==0 else 'KOSDAQ'} API 전수조사 시작...")
    
    for page in range(1, 31):
        params = {
            'sosok': sosok,
            'page': page,
            'field': ['quant', 'open', 'high', 'low', 'frate']
        }
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
            'Referer': 'https://finance.naver.com/sise/sise_market_sum.naver',
            'Accept': '*/*'
        }
        
        try:
            # requests.get으로 직접 표 읽기 시도
            resp = requests.get(base_url, params=params, headers=headers, timeout=15)
            resp.encoding = 'euc-kr'
            
            # read_html 대신 데이터 직접 추출로 안정성 확보
            dfs = pd.read_html(resp.text)
            df = next((d for d in dfs if '종목명' in d.columns), None)
            
            if df is None or df.dropna(subset=['종목명']).empty:
                break
                
            df = df.dropna(subset=['종목명']).copy()
            df.columns = [str(c).strip() for c in df.columns]
            
            for _, row in df.iterrows():
                try:
                    def clean(val): return float(str(val).replace(',', '').replace('%', '').replace('+', '').replace('-', '0'))
                    all_stocks.append({
                        'Name': str(row['종목명']),
                        'Close': int(clean(row['현재가'])),
                        'Ratio': float(clean(row['등락률'])),
                        'Volume': int(clean(row['거래량'])),
                        'Open': int(clean(row['시가'])),
                        'High': int(clean(row['고가'])),
                        'Low': int(clean(row['저가']))
                    })
                except: continue
                
            print(f"✅ {page}p 완료 (누적 {len(all_stocks)}개)")
            time.sleep(random.uniform(0.5, 1.0))
            
        except Exception as e:
            print(f"⚠️ {page}p 오류: {e}")
            break
            
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    
    # 코스피(0) 수집
    df = await fetch_api_data(0)
    
    if df.empty:
        print("❌ 데이터를 가져오는데 실패했습니다.")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    # 필터링 및 정렬
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 지수님 요구사항 디자인 세팅
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
                    v = abs(float(val)) if val else 0
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
    with open(file_name, 'rb') as f:
        await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
