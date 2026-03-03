import os, pandas as pd, asyncio, time
from yahooquery import Ticker
import FinanceDataReader as fdr
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_integrated_data():
    """요구사항 1, 2 반영: 한글명 매칭 및 전종목 수집"""
    print("📡 [1단계] KRX 종목 리스트 확보 및 한글명 매칭 중...")
    try:
        # FDR로 리스트만 가져오는 건 차단 안 됨 (한글명 확보용)
        krx = fdr.StockListing('KRX')[['Code', 'Name', 'Market']]
        # 데이터 누락 방지를 위해 상위 1000개 타겟팅
        target_df = krx.head(1000) 
        kor_name_map = dict(zip(target_df['Code'], target_df['Name']))
        tickers = [c + (".KS" if m == 'KOSPI' else ".KQ") for c, m in zip(target_df['Code'], target_df['Market'])]
    except Exception as e:
        print(f"⚠️ 리스트 확보 실패: {e}")
        return pd.DataFrame()

    all_stocks = []
    chunk_size = 20 # 요구사항 5: 안정적인 수집을 위해 20개씩 끊어서 진행
    
    print(f"📡 [2단계] 야후 엔진으로 {len(tickers)}개 종목 정밀 분석 시작...")
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            t = Ticker(batch, asynchronous=True)
            p_data = t.price
            d_data = t.summary_detail
            
            for symbol in batch:
                p = p_data.get(symbol, {})
                d = d_data.get(symbol, {})
                if isinstance(p, dict) and 'regularMarketPrice' in p:
                    close_p = p.get('regularMarketPrice') or d.get('previousClose') or 0
                    if close_p <= 0: continue
                    
                    code_only = symbol.split('.')[0]
                    all_stocks.append({
                        'Code': code_only,
                        'Name': kor_name_map.get(code_only, code_only), # 한글명 우선 적용
                        'Market': "KOSPI" if symbol.endswith(".KS") else "KOSDAQ",
                        'Open': int(p.get('regularMarketOpen') or d.get('open') or close_p),
                        'Close': int(close_p),
                        'Low': int(p.get('regularMarketDayLow') or d.get('dayLow') or close_p),
                        'High': int(p.get('regularMarketDayHigh') or det.get('dayHigh') or close_p),
                        'Ratio': float(p.get('regularMarketChangePercent', 0) * 100),
                        'Volume': int(p.get('regularMarketVolume') or d.get('volume') or 0)
                    })
        except: continue
        time.sleep(0.1) # 차단 방지

    print(f"✅ [3단계] 수집 완료: {len(all_stocks)}개 종목 확보")
    return pd.DataFrame(all_stocks)

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = await get_integrated_data()
    
    if df.empty: return

    # 요구사항 3: 요일별 모드 (일요일=주간, 나머지=일일)
    is_sun = (now.weekday() == 6)
    report_type = "주간평균" if is_sun else ("일일(금요마감)" if now.weekday() == 5 else "일일")
    file_name = f"{now.strftime('%m%d')}_국내증시_{report_type}.xlsx"
    
    # 요구사항 4: 엑셀 디자인 및 서식
    h_map = {'Code':'종합코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
    f_red, f_ora, f_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    f_head, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                cond_m = df['Market'] == m
                cond_r = (df['Ratio'] >= 5) if trend == '상승' else (df['Ratio'] <= -5)
                
                sub_df = df[cond_m & cond_r].sort_values('Ratio', ascending=(trend=='하락'))
                sub_df = sub_df.drop(columns=['Market']).rename(columns=h_map)
                
                sheet_name = f"{m}_{trend}"
                sub_df.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 헤더 스타일 (짙은 배경 + 흰색 글자)
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border

                # 본문 스타일 (천단위 콤마, 중앙정렬, 등락률 색상)
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
                        if c in [3, 4, 5, 6, 8]: ws.cell(r, c).number_format = '#,##0' # 콤마
                        if c == 7: ws.cell(r, c).number_format = '0.00' # 소수점

                ws.column_dimensions['B'].width = 20 # 종목명 너비 확보

    # 요구사항 5: 텔레그램 전송 (안정화)
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 국내증시 {report_type}\n"
               f"📊 분석: {'주간 변동성' if is_sun else '전수조사'}\n"
               f"📈 상승(5%↑): {len(df[df['Ratio']>=5])}개\n"
               f"📉 하락(5%↓): {len(df[df['Ratio']<=-5])}개")
        with open(file_name, 'rb') as doc:
            await bot.send_document(CHAT_ID, doc, caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
