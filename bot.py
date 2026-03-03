import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font
import time

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        # 1. 전 종목 리스트 확보 (차단 대비 리트라이)
        df_base = None
        for _ in range(3):
            try:
                df_base = fdr.StockListing('KRX')
                if df_base is not None and not df_base.empty: break
            except:
                time.sleep(2)
        
        if df_base is None or df_base.empty: return

        # 2. 요일별 모드 설정 (지수님 요청 로직 정확히 반영)
        if day_of_week == 6: # [일요일] 주간 정밀 분석
            report_type = "주간평균"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
            
            # [기능유지] 시총 상위 500개
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
            sem = asyncio.Semaphore(5) # Expecting value 에러 방지를 위해 접속 수 제한

            async def fetch_weekly(row, semaphore):
                async with semaphore:
                    try:
                        # [오류수정] Yahoo 소스 우회로 서버 차단 방지
                        ticker = f"{row['Code']}.KS" if row['Market'] == 'KOSPI' else f"{row['Code']}.KQ"
                        h = fdr.DataReader(ticker, start_d, end_d)
                        if h is None or len(h) < 2: return None
                        h['rt'] = h['Close'].pct_change() * 100
                        return {
                            'Code': row['Code'], 'Name': row['Name'], 
                            'Open': h.iloc[-1]['Open'], 'Close': h.iloc[-1]['Close'],
                            'Low': h['Low'].min(), 'High': h['High'].max(), 
                            'Ratio': round(h['rt'].mean(), 2), 'Volume': h.iloc[-1]['Volume'],
                            'Market': row['Market']
                        }
                    except: return None

            tasks = [fetch_weekly(row, sem) for _, row in df_target.iterrows()]
            results = await asyncio.gather(*tasks)
            df_final = pd.DataFrame([r for r in results if r is not None])
            target_date_str = f"{start_d}~{end_d}"
            analysis_info = "시가총액 상위 500"

        else: # [화~토] 일일 초고속 분석 (전 종목)
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)" # [복구] 명칭 유지
            target_date_str = now.strftime('%Y-%m-%d')
            
            # [복구] 지수님 원본의 정교한 형변환 로직
            df_base['Close'] = pd.to_numeric(df_base['Close'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_base['Changes'] = pd.to_numeric(df_base['Changes'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            df_base['Volume'] = pd.to_numeric(df_base['Volume'].astype(str).str.replace(',', ''), errors='coerce').fillna(0)
            
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in df_base.columns), None)
            if ratio_col:
                df_base['Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
                if df_base['Ratio'].max() <= 1.0: df_base['Ratio'] *= 100
            else:
                df_base['Ratio'] = (df_base['Changes'] / (df_base['Close'] - df_base['Changes']).replace(0, 1) * 100).fillna(0)
            
            # [복구] 지수님 요청 컬럼 순서
            df_final = df_base[['Code', 'Name', 'Open', 'Close', 'Low', 'High', 'Ratio', 'Volume', 'Market']].copy()
            analysis_info = "전 종목 전수조사"

        if df_final.empty: return

        # 3. 분류 로직 (지수님 원본 시트명 및 필터링 유지)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_sub_market(market_name, is_up):
            temp = df_final[df_final['Market'].str.contains(market_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets_data = {
            '코스피_상승': get_sub_market('KOSPI', True), '코스닥_상승': get_sub_market('KOSDAQ', True),
            '코스피_하락': get_sub_market('KOSPI', False), '코스닥_하락': get_sub_market('KOSDAQ', False)
        }

        # 4. 엑셀 생성 및 디자인 (지수님 원본 디자인 100%)
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    # [복구] 모든 셀 가운데 정렬
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center', vertical='center')
                    
                    # [복구] 등락률 강조 색상 로직
                    ratio_val = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    if ratio_val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif ratio_val >= 20: name_cell.fill = fill_orange
                    elif ratio_val >= 10: name_cell.fill = fill_yellow
                    
                    # [복구] 숫자 콤마 및 소수점 포맷팅
                    for col_idx in [3, 4, 5, 6, 8]: 
                        ws.cell(row, col_idx).number_format = '#,##0'
                    ws.cell(row, 7).number_format = '0.00'

                for i in range(1, 9): ws.column_dimensions[chr(64+i)].width = 15

        # 5. 전송 (지수님 원본 메시지 형식 복구)
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📊 분석기준: {analysis_info}\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 🟡10%↑, 🟠20%↑, 🔴28%↑") # [복구] 가이드 문구 추가
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)

    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
