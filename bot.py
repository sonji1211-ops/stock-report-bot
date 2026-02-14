import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font
import traceback

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() # 5:토, 6:일
    
    # [날짜 보정 로직]
    if day_of_week == 6: # 일요일 실행 (주간 누적)
        report_type = "주간누적(월~금)"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d') # 금요일
        start_date_str = (now - timedelta(days=6)).strftime('%Y-%m-%d')  # 월요일
    elif day_of_week == 5: # 토요일 실행 (금요일 마감 데이터)
        report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1)).strftime('%Y-%m-%d') # 금요일
        start_date_str = target_date_str
    else: # 평일
        report_type = "일일"
        target_date_str = now.strftime('%Y-%m-%d')
        start_date_str = target_date_str

    try:
        # 1. 데이터 수집 (날짜를 명시적으로 지정)
        # StockListing은 최신 상태를 가져오므로, 특정 날짜 데이터를 위해 DataReader와 조합
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty:
            async with bot: await bot.send_message(CHAT_ID, "❌ KRX 종목 리스트를 불러올 수 없습니다.")
            return

        if day_of_week == 6: # 일요일 주간 누적
            weekly_data = []
            df_target = df_base.sort_values(by='Volume', ascending=False).head(800)
            for idx, row in df_target.iterrows():
                try:
                    d_hist = fdr.DataReader(row['Code'], start_date_str, target_date_str)
                    if not d_hist.empty and len(d_hist) >= 2:
                        open_p, close_p = d_hist.iloc[0]['Open'], d_hist.iloc[-1]['Close']
                        ratio = round(((close_p - open_p) / open_p) * 100, 2)
                        weekly_data.append({
                            'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                            'Open': open_p, 'High': d_hist['High'].max(), 'Low': d_hist['Low'].min(),
                            'Close': close_p, 'Calculated_Ratio': ratio, 'Volume': d_hist['Volume'].mean()
                        })
                except: continue
            df = pd.DataFrame(weekly_data)
        else: # 평일 및 토요일 (일일 데이터)
            # 토요일/공휴일 등 장이 안 열리는 날을 대비해 마지막 거래일 데이터를 가져옴
            cols = df_base.columns.tolist()
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), None)
            df_base['Calculated_Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
            if df_base['Calculated_Ratio'].abs().max() < 2: df_base['Calculated_Ratio'] *= 100
            df = df_base.copy()
            df['Calculated_Ratio'] = df['Calculated_Ratio'].round(2)

        # 2. 엑셀 가공 및 색상 (28%↑🔴, 20%↑🟠, 10%↑🟡)
        h_map = {'Code': '종목코드', 'Name': '종목명', 'Market': '시장', 'Open': '시가', 
                 'High': '고가', 'Low': '저가', 'Close': '종가', 'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'}

        def process_data(market, is_up):
            m_df = df[df['Market'].str.contains(market, na=False)].copy()
            res = m_df[m_df['Calculated_Ratio'] >= 5] if is_up else m_df[m_df['Calculated_Ratio'] <= -5]
            res = res.sort_values(by='Calculated_Ratio', ascending=not is_up)
            return res[[c for c in h_map.keys() if c in res.columns]].rename(columns=h_map)

        sheets_data = {'코스피_상승': process_data('KOSPI', True), '코스닥_상승': process_data('KOSDAQ', True),
                       '코스피_하락': process_data('KOSPI', False), '코스닥_하락': process_data('KOSDAQ', False)}

        file_name = f"{now.strftime('%Y-%m-%d')}_국내리포트.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill(start_color="FF0000", fill_type="solid"), PatternFill(start_color="FFCC00", fill_type="solid"), PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row=row, column=8).value or 0))
                    name_cell = ws.cell(row=row, column=2)
                    if val >= 28: name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: name_cell.fill = fill_orange
                    elif val >= 10: name_cell.fill = fill_yellow
                    for c in range(1, 10):
                        cell = ws.cell(row=row, column=c)
                        cell.alignment = Alignment(horizontal='center')
                        if c == 8: cell.number_format = '0.00'
                        elif c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                for i in range(1, 10): ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 15

        # 3. 전송
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 국장 리포트\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"💡 가이드: (🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
    
    except Exception as e:
        async with bot: await bot.send_message(CHAT_ID, f"⚠️ 국장 분석 중 오류: {str(e)}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
