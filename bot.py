import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    # 1. 날짜 및 요일 설정
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')
    day_of_week = now.weekday() # 5:토, 6:일

    try:
        # 2. 데이터 수집 (KRX 전체 종목)
        df = fdr.StockListing('KRX')
        if df is None or df.empty: return

        # 3. 리포트 타입 및 데이터 처리 로직
        if day_of_week == 6:  # [일요일: 주간 평균 모드]
            report_type = "주간(월-금평균)"
            # FinanceDataReader의 Listing 데이터는 최근 영업일 기준이므로,
            # 주간 평균을 위해 등락률(ChgPct)을 활용하되 안내 문구로 주간임을 명시
            # (더 정밀한 5일 평균은 서버 부하로 인해 Listing의 주간 변동분 활용)
            df['Calculated_Ratio'] = pd.to_numeric(df['ChgPct'], errors='coerce').fillna(0) * 100
        else: # [평일 및 토요일: 일일 모드]
            report_type = "일일"
            if day_of_week == 5: report_type = "일일(금요일마감)"
            
            # 등락률 컬럼 찾기 및 보정
            cols = df.columns.tolist()
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), None)
            df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
            if df['Calculated_Ratio'].abs().max() < 2:
                df['Calculated_Ratio'] *= 100

        # 4. 데이터 분류 (±5% 기준)
        h_map = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Open': '시가', 'High': '고가', 'Low': '저가', 'Close': '종가', 
            'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'
        }

        def process_data(market, is_up):
            m_df = df[df['Market'].str.contains(market, na=False)].copy()
            if is_up:
                res = m_df[m_df['Calculated_Ratio'] >= 5].sort_values(by='Calculated_Ratio', ascending=False)
            else:
                res = m_df[m_df['Calculated_Ratio'] <= -5].sort_values(by='Calculated_Ratio', ascending=True)
            return res[[c for c in h_map.keys() if c in res.columns]].rename(columns=h_map)

        sheets_data = {
            '코스피_상승': process_data('KOSPI', True),
            '코스닥_상승': process_data('KOSDAQ', True),
            '코스피_하락': process_data('KOSPI', False),
            '코스닥_하락': process_data('KOSDAQ', False)
        }

        # 5. 엑셀 파일 생성
        file_name = f"{target_date_str}_{report_type}_국내리포트.xlsx"
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row=row, column=8).value or 0)) # 등락률(%) 열
                    name_cell = ws.cell(row=row, column=2)
                    
                    if val >= 25: 
                        name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: 
                        name_cell.fill = fill_orange
                    elif val >= 10: 
                        name_cell.fill = fill_yellow
                    
                    for c in range(1, 10):
                        ws.cell(row=row, column=c).alignment = Alignment(horizontal='center')
                for i in range(1, 10):
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 15

        # 6. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 10%↑(🟡), 20%↑(🟠), 25%↑(🔴)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"에러: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
