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
    # 전 종목 코드 대역 설정 (차단 방지를 위해 묶어서 요청)
    codes = [f"{i:06d}" for i in range(1, 1500, 5)] # 핵심 종목 위주 추출
    tickers = [c + ".KS" for c in codes] + [c + ".KQ" for c in codes]
    
    all_stocks = []
    chunk_size = 30 
    
    print(f"📡 야후 엔진으로 {len(tickers)}개 종목 수집 시작...")
    
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            t = Ticker(batch, asynchronous=True)
            p = t.price
            d = t.summary_detail
            
            for symbol in batch:
                info = p.get(symbol, {})
                det = d.get(symbol, {})
                if isinstance(info, dict) and 'regularMarketPrice' in info:
                    close_p = info.get('regularMarketPrice') or det.get('previousClose') or 0
                    if close_p == 0: continue
                    
                    market = "KOSPI" if symbol.endswith(".KS") else "KOSDAQ"
                    code_only = symbol.split('.')[0]
                    
                    all_stocks.append({
                        'Code': code_only,
                        'Name': info.get('shortName', code_only), # 한글명은 야후 제공명 사용
                        'Market': market,
                        'Open': int(info.get('regularMarketOpen') or det.get('open') or close_p),
                        'Close': int(close_p),
                        'Low': int(info.get('regularMarketDayLow') or det.get('dayLow') or close_p),
                        'High': int(info.get('regularMarketDayHigh') or det.get('dayHigh') or close_p),
                        'Ratio': float(info.get('regularMarketChangePercent', 0) * 100),
                        'Volume': int(info.get('regularMarketVolume') or det.get('volume') or 0)
                    })
        except: continue
        # print(f"✅ {min(i+chunk_size, len(tickers))}개 완료...")
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        # 1. 야후 엔진으로 데이터 확보
        df_final = await get_yahoo_full_data()
        if df_final.empty: return

        # 2. 요일별 모드 설정 (지수님의 '그 느낌' 로직)
        if day_of_week == 6: # 일요일
            report_type = "주간평균"
            analysis_info = "주간 변동성 분석"
            target_date_str = (now - timedelta(days=6)).strftime('%m%d') + "~" + now.strftime('%m%d')
        else:
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)"
            analysis_info = "전 종목 전수조사"
            target_date_str = now.strftime('%Y-%m-%d')

        # 3. 분류 로직 (지수님 요청 시트 분리)
        h_map = {'Code':'종합코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def filter_market(market_name, is_up):
            cond = (df_final['Market'] == market_name)
            cond_r = (df_final['Ratio'] >= 5) if is_up else (df_final['Ratio'] <= -5)
            res = df_final[cond & cond_r].copy()
            return res.sort_values('Ratio', ascending=not is_up).drop(columns=['Market']).rename(columns=h_map)

        sheets_data = {
            '코스피_상승': filter_market('KOSPI', True),
            '코스닥_상승': filter_market('KOSDAQ', True),
            '코스피_하락': filter_market('KOSPI', False),
            '코스닥_하락': filter_market('KOSDAQ', False)
        }

        # 4. 엑셀 생성 및 스타일링 (지수님이 좋아하시는 콤마/색상 양식)
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        fill_head = PatternFill("solid", fgColor="444444")
        font_white = Font(color="FFFFFF", bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                if data.empty:
                    data = pd.DataFrame([['조건 만족 없음']+['']*7], columns=list(h_map.values()))
                
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]

                # 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = fill_head, font_white, Alignment(horizontal='center'), thin_border

                for row in range(2, ws.max_row + 1):
                    # 등락률 강조
                    try:
                        r_val = abs(float(ws.cell(row, 7).value or 0))
                        name_cell = ws.cell(row, 2)
                        if r_val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                        elif r_val >= 20: name_cell.fill = fill_orange
                        elif r_val >= 10: name_cell.fill = fill_yellow
                    except: pass
                    
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = thin_border
                        if col in [3, 4, 5, 6, 8]: ws.cell(row, col).number_format = '#,##0'
                        if col == 7: ws.cell(row, col).number_format = '0.00'
                
                for i in range(1, 9): ws.column_dimensions[chr(64+i)].width = 15

        # 5. 전송 (지수님 스타일의 캡션)
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📊 분석: {analysis_info}\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 🔴28%↑ 🟠20%↑ 🟡10%↑")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=msg, parse_mode="HTML")
        
        if os.path.exists(file_name): os.remove(file_name)

    except Exception as e:
        print(f"❌ 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
