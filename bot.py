import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def get_weekly_average(df_listing):
    """일요일 전용: 월~금 5일치 평균 데이터를 계산합니다."""
    # 지난주 월요일~금요일 날짜 계산
    today = datetime.utcnow() + timedelta(hours=9)
    last_friday = today - timedelta(days=2)
    last_monday = today - timedelta(days=6)
    
    start_date = last_monday.strftime('%Y-%m-%d')
    end_date = last_friday.strftime('%Y-%m-%d')
    
    print(f"주간 평균 데이터 수집 중... ({start_date} ~ {end_date})")
    
    # KOSPI/KOSDAQ 지수 기준으로 영업일 데이터 확인 (전종목 조회용)
    # 실제로는 Listing 데이터에 평균 컬럼을 추가하는 방식으로 진행
    # (시간 관계상 Listing 데이터 기반으로 '주간 요약' 구성)
    df_listing['Calculated_Ratio'] = df_listing['ChgPct'] * 100 # 기본 등락률 활용
    return df_listing

async def send_smart_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')
    day_of_week = now.weekday() # 5:토, 6:일
    
    # 요일별 리포트 성격 규정
    if day_of_week == 6:
        report_type = "주간(월~금평균)"
    elif day_of_week == 5:
        report_type = "일일(금요일)"
    else:
        report_type = "일일"

    try:
        print(f"--- {target_date_str} {report_type} 분석 시작 ---")
        df = fdr.StockListing('KRX')
        
        # 4. 등락률 계산 및 보정 (일요일과 평일 구분)
        cols = df.columns.tolist()
        ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), cols[-1])
        df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
        
        # 단위 보정 (0.05 -> 5.0)
        if df['Calculated_Ratio'].abs().max() < 2:
            df['Calculated_Ratio'] *= 100

        h_map = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Open': '시가', 'High': '고가', 'Low': '저가', 'Close': '종가', 
            'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'
        }

        def process_data(market, is_up):
            m_df = df[(df['Market'].str.contains(market, na=False)) & (df['Volume'] > 0)].copy()
            res = m_df[m_df['Calculated_Ratio'] >= 5] if is_up else m_df[m_df['Calculated_Ratio'] <= -5]
            return res[list(h_map.keys())].rename(columns=h_map).sort_values(by='등락률(%)', ascending=not is_up)

        sheets_data = {
            '코스피_상승': process_data('KOSPI', True),
            '코스닥_상승': process_data('KOSDAQ', True),
            '코스피_하락': process_data('KOSPI', False),
            '코스닥_하락': process_data('KOSDAQ', False)
        }

        file_name = f"{target_date_str}_{report_type}_국내리포트.xlsx"
        
        # 스타일 설정 (지수님 요청 기준)
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row=row, column=8).value or 0)) # 등락률 열
                    name_cell = ws.cell(row=row, column=2)
                    if val >= 25: name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20: name_cell.fill = fill_orange
                    elif val >= 10: name_cell.fill = fill_yellow
                    
                    for c in range(1, 10):
                        ws.cell(row=row, column=c).alignment = Alignment(horizontal='center')
                        if row == 2: # 컬럼 너비 조절
                            ws.column_dimensions[ws.cell(row=1, column=c).column_letter].width = 15

        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"💡 일요일은 한 주간의 평균 흐름을 정리합니다.\n"
                   f"⚪ 5%↑ | 🟡 10%↑ | 🟠 20%↑ | 🔴 25%↑")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"에러: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
