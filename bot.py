import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font
import time

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    if day_of_week == 6: # 일요일: 주간 일별 등락률 평균 모드
        report_type = "주간평균(월~금)"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d')
        start_d, end_d = (now - timedelta(days=6)).strftime('%Y-%m-%d'), target_date_str
        sample_count = 1000 
    else: # 화~토: 일일 리포트
        report_type = "일일"
        if day_of_week == 5: report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1 if day_of_week == 5 else 0)).strftime('%Y-%m-%d')
        start_d = end_d = target_date_str
        sample_count = 0

    try:
        df_base = fdr.StockListing('KRX')
        if df_base is None or df_base.empty: return

        df_target = df_base.sort_values(by='Volume', ascending=False).head(sample_count if sample_count > 0 else len(df_base)).copy()
        res_list = []

        for idx, row in df_target.iterrows():
            try:
                # 데이터 범위 설정 (일일 보고서는 전일 종가가 필요하므로 5일 전부터 조회)
                h = fdr.DataReader(row['Code'], (datetime.strptime(start_d, '%Y-%m-%d') - timedelta(days=5)).strftime('%Y-%m-%d'), end_d)
                
                if not h.empty:
                    if day_of_week == 6: # [일요일] 주간 평균 등락률 계산
                        weekly_data = h.loc[start_d:end_d].copy()
                        if len(weekly_data) >= 2:
                            # 매일의 등락률(종가 기준) 계산 후 평균 산출
                            weekly_data['daily_rt'] = weekly_data['Close'].pct_change() * 100
                            avg_ratio = round(weekly_data['daily_rt'].mean(), 2)
                            
                            res_list.append({
                                'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                                'Open': weekly_data.iloc[-1]['Open'], 'High': weekly_data['High'].max(),
                                'Low': weekly_data['Low'].min(), 'Close': weekly_data.iloc[-1]['Close'],
                                'Calculated_Ratio': avg_ratio, 'Volume': weekly_data.iloc[-1]['Volume']
                            })
                    else: # [평일/토요일] 일일 등락률 계산
                        if len(h) >= 2:
                            o, c = h.iloc[-2]['Close'], h.iloc[-1]['Close']
                            ratio = round(((c - o) / o) * 100, 2)
                            res_list.append({
                                'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                                'Open': h.iloc[-1]['Open'], 'High': h.iloc[-1]['High'],
                                'Low': h.iloc[-1]['Low'], 'Close': c,
                                'Calculated_Ratio': ratio, 'Volume': h.iloc[-1]['Volume']
                            })
            except: continue
            if idx % 300 == 0: await asyncio.sleep(0.1)

        df_final = pd.DataFrame(res_list)
        if df_final.empty: return

        # 분류 및 엑셀 작업
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Market':'시장', 'Open':'시가', 'High':'고가', 'Low':'저가', 'Close':'종가', 'Calculated_Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_sub(market, is_up):
            m_df = df_final[df_final['Market'].str.contains(market, na=False)].copy()
            # 평균값이므로 기준을 5%에서 2%로 낮출지 고민해보세요. 일단 요청대로 5% 유지합니다.
            cond = (m_df['Calculated_Ratio'] >= 5) if is_up else (m_df['Calculated_Ratio'] <= -5)
            return m_df[cond].sort_values('Calculated_Ratio', ascending=not is_up)[list(h_map.keys())].rename(columns=h_map)

        sheets_data = {'코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True), '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)}

        file_name = f"{now.strftime('%Y-%m-%d')}_국내리포트.xlsx"
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
                        elif c in [4, 5, 6, 7, 9]: ws.cell(row, c).number_format = '#,##0'
                for i in range(1, 10): ws.column_dimensions[ws.cell(1, i).column_letter].width = 15

        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀에서 종목명 색깔을 확인하세요!\n(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f: await bot.send_document(CHAT_ID, f, caption=msg)

    except Exception as e: print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
