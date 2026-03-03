import os, pandas as pd, requests, asyncio
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def fetch_real_data(sosok):
    # KOSPI: KOSPI, KOSDAQ: KOSDAQ
    market = "KOSPI" if sosok == 0 else "KOSDAQ"
    all_stocks = []
    
    print(f"📡 {market} 모바일 엔진 가동 (전수조사 시작)...")
    
    # 네이버 모바일 증권 실제 데이터 API 주소
    # pageSize를 100으로 올려서 요청 횟수를 줄여 차단 원천 방지
    for page in range(1, 26): # 100개씩 25페이지면 2,500개 (상장사 전수조사)
        url = f"https://m.stock.naver.com/api/stock/marketValue/{market}?page={page}&pageSize=100"
        headers = {
            'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
            'Referer': 'https://m.stock.naver.com/'
        }
        
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            data = resp.json()
            stocks = data.get('stocks', [])
            
            if not stocks: break
            
            for s in stocks:
                # API에서 주는 순수 숫자 데이터 바로 매핑
                all_stocks.append({
                    'Name': s['stockName'],
                    'Close': int(s['closePrice'].replace(',', '')),
                    'Ratio': float(s['fluctuationsRatio']),
                    'Volume': int(s['accumulatedTradingVolume']),
                    'Open': int(s['openPrice'].replace(',', '')),
                    'High': int(s['highPrice'].replace(',', '')),
                    'Low': int(s['lowPrice'].replace(',', ''))
                })
            
            if page % 5 == 0: print(f"✅ {page*100}위까지 수집 완료...")
            
        except Exception as e:
            print(f"⚠️ {page}p 오류: {e}")
            break
            
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    
    # 0은 코스피, 1은 코스닥 (파일명에 맞춰 수정하세요)
    # KOSPI용 파일이면 0, KOSDAQ용 파일이면 1
    df = await fetch_real_data(0) 
    
    if df.empty:
        print("❌ 모바일 엔진으로도 데이터를 가져오지 못했습니다.")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 지수님 요구사항 디자인 (100% 동일 적용)
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

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📊 전수조사: {len(df)}개 완료\n📈 상승: {len(up_df)} / 📉 하락: {len(down_df)}\n💡 🔴28% 🟠20% 🟡10%"
    with open(file_name, 'rb') as f:
        await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
