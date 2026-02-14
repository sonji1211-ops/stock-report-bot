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

    # 1. 날짜 및 타입 설정
    if day_of_week == 6: # 일요일: 주간 누적
        report_type = "주간누적(월~금)"
        target_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d') # 금요일
        start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d') # 월요일
        end_d = target_date_str
    else: # 화~토: 일일
        report_type = "일일"
        if day_of_week == 5: report_type = "일일(금요일마감)"
        target_date_str = (now - timedelta(days=1 if day_of_week == 5 else 0)).strftime('%Y-%m-%d')
        start_d = end_d = target_date_str

    try:
        # 2. 데이터 수집 (예비 경로 포함 5회 재시도)
        df_base = None
        for i in range(5):
            try:
                # 첫 번째 시도: KRX 전체 목록
                df_base = fdr.StockListing('KRX')
                if df_base is not None and not df_base.empty: break
                # 실패 시 예비: KOSPI/KOSDAQ 각각 시도
                df_base = pd.concat([fdr.StockListing('KOSPI'), fdr.StockListing('KOSDAQ')])
                if not df_base.empty: break
            except:
                time.sleep(5) # 5초 대기 후 재시도
        
        if df_base is None or df_base.empty:
            async with bot: await bot.send_message(CHAT_ID, "❌ [국장] 거래소 서버 응답 없음. (주말 점검 중)")
            return

        # 3. 데이터 가공
        if day_of_week == 6:
            # 주간 누적: 상위 1,000개
            df_target = df_base.sort_values(by='Volume', ascending=False).head(1000).copy()
            res_list = []
            for idx, row in df_target.iterrows():
                try:
                    h = fdr.DataReader(row['Code'], start_d, end_d)
                    if len(h) >= 2:
                        o, c = h.iloc[0]['Open'], h.iloc[-1]['Close']
                        ratio = round(((c - o) / o) * 100, 2)
                        res_list.append({
                            'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                            'Open': o, 'High': h['High'].max(), 'Low': h['Low'].min(),
                            'Close': c, 'Calculated_Ratio': ratio, 'Volume': h['Volume'].mean()
                        })
                    time.sleep(0.05) # 서버 차단 방지용 미세 딜레이
                except: continue
            df = pd.DataFrame(res_list)
        else:
            # 일일 리포트
            df = df_base.copy()
            # 서버가 등락률(ChgPct)을 0으로 줬을 경우 대비 수동 계산 로직 포함
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in df.columns), None)
            df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
            if df['Calculated_Ratio'].abs().max() < 2: df['Calculated_Ratio'] *= 100
            df['Calculated_Ratio'] = df['Calculated_Ratio'].round(2)

        # 4. 엑셀 분류 및 생성 (지수님 커스텀 스타일)
        h_map = {'Code': '종목코드', 'Name': '종목명', 'Market': '시장', 'Open': '시가', 
                 'High': '고가', 'Low': '저가', 'Close': '종가', 'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'}

        def process_data(market, is_up):
            m_df = df[df['Market'].str.contains(market, na=False)].copy()
            if is_up: return m_df[m_df['Calculated_Ratio'] >= 5].sort_values('Calculated_Ratio', ascending=False)[list(h_map.keys())].rename(columns=h_map)
            return m_df[m_df['Calculated_Ratio'] <= -5].sort_values('Calculated_Ratio', ascending=True)[list(h_map.keys())].rename(columns=h_map)

        sheets_data = {'코스피_상승': process_data('KOSPI', True), '코스닥_상승': process_data('KOSDAQ', True),
                       '코스피_하락': process_data('KOSPI', False), '코스닥_하락': process_data('KOSDAQ', False)}

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

        # 5. 전송 (지수님 전용 메시지 포맷)
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀에서 종목명 색깔을 확인하세요!\n"
                   f"(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
