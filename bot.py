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
        if day_of_week == 6: # [일요일] 시총 상위 500개 주간 정밀 분석
            report_type = "주간평균"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d') # 금요일
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d') # 월요일
            
            # 시가총액 순 정렬 후 상위 500개 추출
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
            
            async def fetch_weekly(row):
                try:
                    h = fdr.DataReader(row['Code'], start_d, end_d)
                    if len(h) < 2: return None
                    h['rt'] = h['Close'].pct_change() * 100
                    avg_ratio = round(h['rt'].mean(), 2)
                    return {
                        'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                        'Open': h.iloc[-1]['Open'], 'High': h['High'].max(), 'Low': h['Low'].min(),
                        'Close': h.iloc[-1]['Close'], 'Ratio': avg_ratio, 'Volume': h.iloc[-1]['Volume']
                    }
                except: return None

            # --- 지수님이 말씀하신 병렬 처리 핵심 로직 (일요일 전용) ---
            print("--- 일요일 주간 데이터 병렬 분석 중... ---")
            tasks = [fetch_weekly(row) for _, row in df_target.iterrows()]
            results = await asyncio.gather(*tasks)
            df_final = pd.DataFrame([r for r in results if r is not None])
            # ------------------------------------------------------
            
            target_date_str = f"{start_d}~{end_d}"
            analysis_info = "시가총액 상위 500"

        else: # [화~토] 전 종목 초고속 일일 분석
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)"
            target_date_str = now.strftime('%Y-%m-%d')
            
            # 일일 보고는 이미 df_base에 데이터가 다 있어서 병렬 처리가 필요 없이 바로 변환 (초고속)
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in df_base.columns), 'ChangesRatio')
            df_base['Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
            df_final = df_base[['Code', 'Name', 'Market', 'Open', 'High', 'Low', 'Close', 'Ratio', 'Volume']].copy()
            analysis_info = "전 종목 전수조사"

        if df_final is None or df_final.empty: return

        # 3. 분류 로직 (상승/하락 5% 기준)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Market':'시장', 'Open':'시가', 'High':'고가', 'Low':'저가', 'Close':'종가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_sub(market, is_up):
            m_df = df_final[df_final['Market'].str.contains(market, na=False)].copy()
            cond = (m_df['Ratio'] >= 5) if is_up else (m_df['Ratio'] <= -5)
            return m_df[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map)

        sheets_data = {'코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True),
                       '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)}

        # 4. 엑셀 생성 및 스타일링
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFCC00"), PatternFill("solid", fgColor="FFFF00")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row, 8).value or 0))
                    name_cell = ws.cell(row, 2)
                    if val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: name_cell.fill = fill_orange
                    elif val >= 10: name_cell.fill = fill_yellow
                    for c in range(1, 10):
                        ws.cell(row, c).alignment = Alignment(horizontal='center')
                        if c == 8: ws.cell(row, c).number_format = '0.00'
                for i in range(1, 10): ws.column_dimensions[chr(64+i)].width = 15

        # 5. 텔레그램 발송
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📊 분석기준: {analysis_info}\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 🟡10%↑, 🟠20%↑, 🔴28%↑")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=msg)

    except Exception as e: print(f"오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
