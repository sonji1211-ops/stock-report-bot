import os, pandas as pd, asyncio, time
from yahooquery import Ticker
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_yahoo_full_data():
    """지수님이 원하시는 전 종목 데이터를 야후에서 긁어오는 엔진"""
    codes = [f"{i:06d}" for i in range(1, 1500, 4)] # 간격을 조절하여 더 많은 종목 확보
    tickers = [c + ".KS" for c in codes] + [c + ".KQ" for c in codes]
    
    all_stocks = []
    chunk_size = 30 
    print(f"📡 야후 엔진으로 {len(tickers)}개 종목 수집 시작...")
    
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            t = Ticker(batch, asynchronous=True)
            p, d = t.price, t.summary_detail
            for symbol in batch:
                info, det = p.get(symbol, {}), d.get(symbol, {})
                if isinstance(info, dict) and 'regularMarketPrice' in info:
                    cp = info.get('regularMarketPrice') or det.get('previousClose') or 0
                    if cp == 0: continue
                    all_stocks.append({
                        'Code': symbol.split('.')[0],
                        'Name': info.get('shortName', symbol.split('.')[0]),
                        'Market': "KOSPI" if symbol.endswith(".KS") else "KOSDAQ",
                        'Open': int(info.get('regularMarketOpen') or det.get('open') or cp),
                        'Close': int(cp),
                        'Low': int(info.get('regularMarketDayLow') or det.get('dayLow') or cp),
                        'High': int(info.get('regularMarketDayHigh') or det.get('dayHigh') or cp),
                        'Ratio': float(info.get('regularMarketChangePercent', 0) * 100),
                        'Volume': int(info.get('regularMarketVolume') or det.get('volume') or 0)
                    })
        except: continue
        time.sleep(0.1) # 깃허브 차단 방지용 미세 지연
    
    print(f"✅ 수집 완료: {len(all_stocks)}개 확보")
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        df_final = await get_yahoo_full_data()
        if df_final.empty:
            print("🚨 데이터 수집 실패")
            return

        # 요일 로직 (지수님 스타일)
        is_sun = (day_of_week == 6)
        report_type = "주간평균" if is_sun else ("일일(금요마감)" if day_of_week == 5 else "일일")
        analysis_info = "주간 변동성 분석" if is_sun else "전 종목 전수조사"
        target_date_str = now.strftime('%Y-%m-%d')
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"

        # 시트 분류 및 서식 설정
        h_map = {'Code':'종합코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        f_red, f_ora, f_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
        f_head, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for m in ['KOSPI', 'KOSDAQ']:
                for t in ['상승', '하락']:
                    sub = df_final[(df_final['Market']==m) & ((df_final['Ratio']>=5) if t=='상승' else (df_final['Ratio']<=-5))]
                    sub = sub.sort_values('Ratio', ascending=(t=='하락')).drop(columns=['Market']).rename(columns=h_map)
                    s_name = f"{m}_{t}"
                    sub.to_excel(writer, sheet_name=s_name, index=False)
                    ws = writer.sheets[s_name]
                    for cell in ws[1]: cell.fill, cell.font, cell.alignment, cell.border = f_head, f_white, Alignment(horizontal='center'), border
                    for r in range(2, ws.max_row + 1):
                        try:
                            rv = abs(float(ws.cell(r, 7).value or 0))
                            if rv >= 28: ws.cell(r, 2).fill, ws.cell(r, 2).font = f_red, f_white
                            elif rv >= 20: ws.cell(r, 2).fill = f_ora
                            elif rv >= 10: ws.cell(r, 2).fill = f_yel
                        except: pass
                        for c in range(1, 9):
                            ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                            if c in [3,4,5,6,8]: ws.cell(r, c).number_format = '#,##0'
                            if c == 7: ws.cell(r, c).number_format = '0.00'
                    ws.column_dimensions['B'].width = 18

        # 📤 전송 (전송 보강: async with 블록 활용)
        print("📤 리포트 전송 중...")
        msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
               f"📊 분석: {analysis_info}\n"
               f"📈 상승(5%↑): {len(df_final[df_final['Ratio']>=5])}개\n"
               f"📉 하락(5%↓): {len(df_final[df_final['Ratio']<=-5])}개\n\n"
               f"💡 🔴28%↑ 🟠20%↑ 🟡10%↑")
        
        async with bot:
            with open(file_name, 'rb') as doc:
                await bot.send_document(CHAT_ID, doc, caption=msg)
        print("✅ 전송 완료")

    except Exception as e:
        print(f"❌ 오류 발생: {e}")
    finally:
        if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(send_smart_report())
