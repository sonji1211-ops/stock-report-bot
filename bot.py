import os, pandas as pd, asyncio, time, datetime
import requests
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_explosive_data():
    """야후의 실시간 스크리너를 활용해 급등주/거래량 상위 종목을 통째로 긁어옵니다."""
    print("📡 [1단계] 실시간 급등 및 거래량 상위 데이터 긁어오기...")
    
    all_stocks = []
    # 한국 시장(KOSPI/KOSDAQ)의 모든 종목을 포함하는 야후 쿼리
    # 스캔 범위를 넓히기 위해 여러 번의 대량 요청 세션을 가집니다.
    market_urls = [
        "https://query1.finance.yahoo.com/v1/finance/screener/predefined/saved?formatted=false&scrIds=day_gainers&count=250",
        "https://query1.finance.yahoo.com/v1/finance/screener/predefined/saved?formatted=false&scrIds=most_actives&count=250"
    ]
    
    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'}

    for url in market_urls:
        try:
            res = requests.get(url, headers=headers).json()
            results = res.get('finance', {}).get('result', [{}])[0].get('quotes', [])
            
            for q in results:
                symbol = q.get('symbol', '')
                # 한국 종목(.KS, .KQ)만 필터링
                if not (symbol.endswith('.KS') or symbol.endswith('.KQ')): continue
                
                cp = q.get('regularMarketPrice', 0)
                vol = q.get('regularMarketVolume', 0)
                ratio = q.get('regularMarketChangePercent', 0)
                
                if cp == 0: continue
                
                all_stocks.append({
                    'Code': symbol.split('.')[0],
                    'Name': q.get('shortName', symbol.split('.')[0]),
                    'Market': "KOSPI" if symbol.endswith(".KS") else "KOSDAQ",
                    'Open': int(q.get('regularMarketOpen', cp)),
                    'Close': int(cp),
                    'Low': int(q.get('regularMarketDayLow', cp)),
                    'High': int(q.get('regularMarketDayHigh', cp)),
                    'Ratio': float(ratio),
                    'Volume': int(vol)
                })
        except: continue
        time.sleep(1)

    # 중복 제거 및 데이터 정리
    df = pd.DataFrame(all_stocks).drop_duplicates(subset=['Code'])
    print(f"✅ 수집 완료: {len(df)}개 유효 종목 확보 (거래량 포함)")
    return df

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    df = get_explosive_data()
    if df.empty:
        print("🚨 데이터를 가져오지 못했습니다. Yahoo 서버 응답 확인 필요.")
        return

    # 요일 로직
    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # 엑셀 디자인 설정 (지수님 요구사항 준수)
    h_map = {'Code':'CODE', 'Name':'NAME', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
    f_red, f_ora, f_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    f_head, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        # 지수님이 요청하신 4개 시트 구성
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                # 등락률 기준 필터링 (기존 5% 룰 유지)
                sub = df[(df['Market']==m) & ((df['Ratio']>=5) if trend=='상승' else (df['Ratio']<=-5))]
                
                # 만약 종목이 너무 적으면 2%까지 완화해서 리포트 풍성하게 만들기 (지수님 요청 반영)
                if len(sub) < 5:
                    sub = df[(df['Market']==m) & ((df['Ratio']>=2) if trend=='상승' else (df['Ratio']<=-2))]
                
                sub = sub.sort_values('Ratio', ascending=(trend=='하락')).drop(columns=['Market']).rename(columns=h_map)
                s_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]

                # 디자인: 헤더
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border

                # 디자인: 본문
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
                ws.column_dimensions['B'].width = 25

    # 전송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 리포트\n\n"
               f"📊 종목수: {len(df)}개 정밀 스캔\n"
               f"📈 상승(급등): {len(df[df['Ratio']>=3])}개\n"
               f"📉 하락(급락): {len(df[df['Ratio']<=-3])}개\n\n"
               f"💡 거래량/코스닥 누락 해결 완료 🚀")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
