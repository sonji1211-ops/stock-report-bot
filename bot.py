import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() # 0:월, ..., 5:토, 6:일

    try:
        # 1. 전 종목 기본 데이터 확보
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        # 2. 요일별 모드 설정
        if day_of_week == 6: # [일요일] 주간 평균 분석 (월~금 데이터)
            report_type = "주간평균(5%↑↓)"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d') # 금요일
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d') # 월요일
            
            # 주간은 시총 상위 500~1000개 정도만 정밀 분석 (속도 및 정확도)
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(1000).copy()
            
            async def fetch_weekly(row):
                try:
                    h = fdr.DataReader(row['Code'], start_d, end_d)
                    if len(h) < 2: return None
                    # 주간 등락률: (금요일 종가 - 월요일 시가) / 월요일 시가 * 100
                    weekly_rt = ((h.iloc[-1]['Close'] - h.iloc[0]['Open']) / h.iloc[0]['Open']) * 100
                    return {
                        'Market': row['Market'], 'Code': row['Code'], 'Name': row['Name'], 
                        'Open': h.iloc[0]['Open'], 'Close': h.iloc[-1]['Close'],
                        'Low': h['Low'].min(), 'High': h['High'].max(), 
                        'Ratio': round(weekly_rt, 2), 'Volume': int(h['Volume'].mean()) # 주간 평균 거래량
                    }
                except: return None

            tasks = [fetch_weekly(row) for _, row in df_target.iterrows()]
            results = await asyncio.gather(*tasks)
            df_final = pd.DataFrame([r for r in results if r is not None])
            target_date_str = f"{start_d} ~ {end_d}"
            analysis_info = "시총 상위 1000개 주간 분석"

        else: # [화~토] 일일 전수조사 (전일자 데이터)
            report_type = "일일"
            target_date_str = (now - timedelta(days=1)).strftime('%Y-%m-%d') if day_of_week == 0 else now.strftime('%Y-%m-%d')
            
            df_base['Ratio'] = pd.to_numeric(df_base['ChgPct'], errors='coerce').fillna(0)
            df_final = df_base[['Market', 'Code', 'Name', 'Open', 'Close', 'Low', 'High', 'Ratio', 'Volume']].copy()
            analysis_info = "국내 전 종목 실시간/전일 분석"

        if df_final.empty: return

        # 3. 분류 및 시트 데이터 준비
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_sheet_data(market, is_up):
            cond_m = df_final['Market'].str.contains(market, na=False)
            cond_r = (df_final['Ratio'] >= 5) if is_up else (df_final['Ratio'] <= -5)
            res = df_final[cond_m & cond_r].copy()
            # 하락은 많이 떨어진 순(-30%가 맨 위)으로 정렬
            return res.sort_values('Ratio', ascending=not is_up).drop(columns=['Market']).rename(columns=h_map)

        sheets_dict = {
            '코스피_상승': get_sheet_data('KOSPI', True),
            '코스피_하락': get_sheet_data('KOSPI', False),
            '코스닥_상승': get_sheet_data('KOSDAQ', True),
            '코스닥_하락': get_sheet_data('KOSDAQ', False)
        }

        # 4. 엑셀 생성 및 디자인
        file_name = f"{now.strftime('%m%d')}_{report_type}_리포트.xlsx"
        fill_red = PatternFill("solid", fgColor="FF0000")    # 25%↑
        fill_orange = PatternFill("solid", fgColor="FFCC00") # 20%↑
        fill_yellow = PatternFill("solid", fgColor="FFFF00") # 10%↑
        header_fill = PatternFill("solid", fgColor="444444")
        font_white = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_dict.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 열 너비 설정
                ws.column_dimensions['B'].width = 25 # 종목명
                for col in ['C', 'D', 'E', 'F', 'H']: ws.column_dimensions[ws.cell(1, data.columns.get_loc(h_map[col if col != 'Volume' else 'Volume'])+1).column_letter].width = 15

                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        cell = ws.cell(r, c)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = border
                        
                        if r == 1: # 헤더
                            cell.fill, cell.font = header_fill, font_white
                        else:
                            # 숫자 포맷팅
                            if c in [3, 4, 5, 6, 8]: cell.number_format = '#,##0'
                            if c == 7: cell.number_format = '0.00'
                            
                            # 색상 강조 로직 (G열 = 등락률)
                            ratio_val = abs(float(ws.cell(r, 7).value or 0))
                            if ratio_val >= 25: ws.cell(r, 2).fill, ws.cell(r, 2).font = fill_red, font_white
                            elif ratio_val >= 20: ws.cell(r, 2).fill = fill_orange
                            elif ratio_val >= 10: ws.cell(r, 2).fill = fill_yellow

        # 5. 텔레그램 전송
        total_up = len(sheets_dict['코스피_상승']) + len(sheets_dict['코스닥_상승'])
        total_down = len(sheets_dict['코스피_하락']) + len(sheets_dict['코스닥_하락'])
        
        async with bot:
            msg = (f"📅 {target_date_str} {report_type}\n"
                   f"━━━━━━━━━━━━━━━\n"
                   f"📊 분석기준: {analysis_info}\n"
                   f"📈 급상승(5%↑): {total_up}개\n"
                   f"📉 급하락(5%↓): {total_down}개\n"
                   f"━━━━━━━━━━━━━━━\n"
                   f"💡 가독성 가이드\n"
                   f"🟡 10%↑ | 🟠 20%↑ | 🔴 25%↑")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
        
        if os.path.exists(file_name): os.remove(file_name)

    except Exception as e:
        print(f"🚨 오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
