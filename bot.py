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
        # 1. 전 종목 기본 데이터 확보
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        # 2. 요일별 모드 설정
        if day_of_week == 6: # [일요일] 주간 정밀 분석
            report_type = "주간평균"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
            
            async def fetch_weekly(row):
                try:
                    h = fdr.DataReader(row['Code'], start_d, end_d)
                    if len(h) < 2: return None
                    h['rt'] = h['Close'].pct_change() * 100
                    return {
                        'Code': row['Code'], 'Name': row['Name'], 
                        'Open': h.iloc[-1]['Open'], 'Close': h.iloc[-1]['Close'],
                        'Low': h['Low'].min(), 'High': h['High'].max(), 
                        'Ratio': round(h['rt'].mean(), 2), 'Volume': h.iloc[-1]['Volume']
                    }
                except: return None

            tasks = [fetch_weekly(row) for _, row in df_target.iterrows()]
            results = await asyncio.gather(*tasks)
            df_final = pd.DataFrame([r for r in results if r is not None])
            target_date_str = f"{start_d}~{end_d}"
            analysis_info = "시가총액 상위 500"

        else: # [화~토] 일일 초고속 분석
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)"
            target_date_str = now.strftime('%Y-%m-%d')
            
            # 수치 형변환
            df_base['Close'] = pd.to_numeric(df_base['Close'], errors='coerce').fillna(0)
            df_base['Changes'] = pd.to_numeric(df_base['Changes'], errors='coerce').fillna(0)
            df_base['Volume'] = pd.to_numeric(df_base['Volume'], errors='coerce').fillna(0)
            
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in df_base.columns), None)
            if ratio_col:
                df_base['Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
            else:
                df_base['Ratio'] = (df_base['Changes'] / (df_base['Close'] - df_base['Changes']) * 100).fillna(0)
            
            # 요청 순서: 종목코드, 종목명, 시가, 종가, 저가, 고가, 등락률, 거래량
            df_final = df_base[['Code', 'Name', 'Open', 'Close', 'Low', 'High', 'Ratio', 'Volume']].copy()
            analysis_info = "전 종목 전수조사"

        if df_final is None or df_final.empty: return

        # 3. 분류 로직
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_sub_market(market, is_up):
            temp_df = df_base[df_base['Market'].str.contains(market, na=False)].copy()
            temp_df['Ratio'] = df_final['Ratio'] 
            res = temp_df[['Code', 'Name', 'Open', 'Close', 'Low', 'High', 'Ratio', 'Volume']].copy()
            cond = (res['Ratio'] >= 5) if is_up else (res['Ratio'] <= -5)
            return res[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map)

        sheets_data = {
            '코스피_상승': get_sub_market('KOSPI', True), '코스닥_상승': get_sub_market('KOSDAQ', True),
            '코스피_하락': get_sub_market('KOSPI', False), '코스닥_하락': get_sub_market('KOSDAQ', False)
        }

        # 4. 엑셀 생성 및 숫자 포맷팅(콤마)
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    # 등락률 강조 (G열 = 7번째)
                    ratio_val = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    if ratio_val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif ratio_val >= 20: name_cell.fill = fill_orange
                    elif ratio_val >= 10: name_cell.fill = fill_yellow
                    
                    # 숫자 포맷팅 (콤마)
                    # 시가(C), 종가(D), 저가(E), 고가(F), 거래량(H) 전체 적용
                    for col_idx in [3, 4, 5, 6, 8]:
                        ws.cell(row, col_idx).number_format = '#,##0'
                    
                    # 등락률(G)은 소수점 2자리
                    ws.cell(row, 7).number_format = '0.00'
                    
                    for c in range(1, 9):
                        ws.cell(row, c).alignment = Alignment(horizontal='center')
                
                for i in range(1, 9): 
                    ws.column_dimensions[chr(64+i)].width = 15

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
