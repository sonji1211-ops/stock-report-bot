import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        # 1. 전 종목 기본 데이터 확보 (서버 차단 시 리트라이)
        try:
            df_base = fdr.StockListing('KRX')
        except:
            # KRX 전체가 에러나면 코스피/코스닥 각각 시도
            df_base = pd.concat([fdr.StockListing('KOSPI'), fdr.StockListing('KOSDAQ')])
            
        if df_base is None or df_base.empty: return

        # 2. 요일별 모드 설정
        if day_of_week == 6: # [일요일] 주간 정밀 분석
            report_type = "주간평균"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
            # 지수님이 요청하신 시총 상위 500개
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
            
            async def fetch_weekly(row):
                try:
                    # 야후 티커 변환 (.KS 또는 .KQ) - Expecting value 에러 방지용
                    ticker = f"{row['Code']}.KS" if row['Market'] == 'KOSPI' else f"{row['Code']}.KQ"
                    h = fdr.DataReader(ticker, start_d, end_d)
                    if h is None or len(h) < 2: return None
                    h['rt'] = h['Close'].pct_change() * 100
                    return {
                        'Code': row['Code'], 'Name': row['Name'], 
                        'Open': h.iloc[-1]['Open'], 'Close': h.iloc[-1]['Close'],
                        'Low': h['Low'].min(), 'High': h['High'].max(), 
                        'Ratio': round(h['rt'].mean(), 2), 'Volume': h.iloc[-1]['Volume'],
                        'Market': row['Market'] # 분류를 위해 유지
                    }
                except: return None

            # [중요] 50개씩 끊어서 수집 (서버 차단 방지)
            results = []
            row_list = [row for _, row in df_target.iterrows()]
            for i in range(0, len(row_list), 50):
                chunk = row_list[i:i+50]
                tasks = [fetch_weekly(r) for r in chunk]
                chunk_res = await asyncio.gather(*tasks)
                results.extend([res for res in chunk_res if res is not None])
                await asyncio.sleep(1) # 1초 휴식

            df_final = pd.DataFrame(results)
            target_date_str = f"{start_d}~{end_d}"
            analysis_info = "시가총액 상위 500"

        else: # [화~토] 일일 초고속 분석
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)"
            target_date_str = now.strftime('%Y-%m-%d')
            
            # 수치 형변환
            for col in ['Close', 'Changes', 'Volume', 'Open', 'Low', 'High']:
                df_base[col] = pd.to_numeric(df_base[col].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in df_base.columns), None)
            if ratio_col:
                df_base['Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
                if df_base['Ratio'].max() <= 1.0: df_base['Ratio'] *= 100 # 소수점 형태일 경우 보정
            else:
                df_base['Ratio'] = (df_base['Changes'] / (df_base['Close'] - df_base['Changes']) * 100).fillna(0)
            
            df_final = df_base[['Code', 'Name', 'Open', 'Close', 'Low', 'High', 'Ratio', 'Volume', 'Market']].copy()
            analysis_info = "전 종목 전수조사"

        if df_final.empty: return

        # 3. 분류 로직
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_sub_market(market_name, is_up):
            temp = df_final[df_final['Market'].str.contains(market_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets_data = {
            '코스피_상승': get_sub_market('KOSPI', True), '코스닥_상승': get_sub_market('KOSDAQ', True),
            '코스피_하락': get_sub_market('KOSPI', False), '코스닥_하락': get_sub_market('KOSDAQ', False)
        }

        # 4. 엑셀 생성 및 디자인
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    # 전체 가운데 정렬
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
                    
                    # 등락률 강조
                    ratio_val = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    if ratio_val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif ratio_val >= 20: name_cell.fill = fill_orange
                    elif ratio_val >= 10: name_cell.fill = fill_yellow
                    
                    # 숫자 포맷
                    for col_idx in [3, 4, 5, 6, 8]: ws.cell(row, col_idx).number_format = '#,##0'
                    ws.cell(row, 7).number_format = '0.00'
                
                for i in range(1, 9): ws.column_dimensions[chr(64+i)].width = 15

        # 5. 전송
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📊 분석기준: {analysis_info}\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 🟡10%↑, 🟠20%↑, 🔴28%↑")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)

    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
