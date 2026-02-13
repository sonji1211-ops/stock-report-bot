import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 직접 입력 모드
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')
    day_of_week = now.weekday() 
    
    if day_of_week == 6:
        print("오늘은 일요일이므로 리포트를 생성하지 않습니다.")
        return

    report_type = "주간" if day_of_week == 5 else "일일"

    try:
        print(f"--- {target_date_str} {report_type} 분석 시작 ---")
        
        df = fdr.StockListing('KRX')
        if df is None or df.empty:
            print("데이터를 가져오는 데 실패했습니다.")
            return

        cols = df.columns.tolist()
        cap_col = next((c for c in ['Marcap', 'Amount', 'MarketCap'] if c in cols), cols[-1])
        chg_amt_col = next((c for c in ['Change', 'Changes', 'ChgAmt'] if c in cols), None)

        # 1. 고가(High), 저가(Low)를 수집 목록에 추가
        needed_cols = ['Open', 'High', 'Low', 'Close', 'Volume', cap_col]
        if chg_amt_col: needed_cols.append(chg_amt_col)
        
        for c in needed_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        if chg_amt_col:
            def calculate_ratio(row):
                prev_close = row['Close'] - row[chg_amt_col]
                return (row[chg_amt_col] / prev_close * 100) if prev_close != 0 else 0
            df['Calculated_Ratio'] = df.apply(calculate_ratio, axis=1)
        else:
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), cols[-1])
            df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
            
        if df['Calculated_Ratio'].abs().max() < 2 and df['Calculated_Ratio'].abs().max() > 0:
            df['Calculated_Ratio'] *= 100

        # 2. 엑셀에 표시될 한글 이름 매핑 (고가, 저가 추가)
        h_map = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Open': '시가', 'High': '고가', 'Low': '저가', 'Close': '종가(현재가)', 
            'Calculated_Ratio': '전일대비(%)', 'Volume': '거래량'
        }

        def process_data(market, is_up):
            m_df = df[(df['Market'].str.contains(market, na=False)) & (df['Volume'] > 0)].copy()
            if is_up:
                res = m_df[m_df['Calculated_Ratio'] >= 5].copy()
                res = res.sort_values(by='Calculated_Ratio', ascending=False)
            else:
                res = m_df[m_df['Calculated_Ratio'] <= -5].copy()
                res = res.sort_values(by='Calculated_Ratio', ascending=True)
            
            actual_cols = [c for c in h_map.keys() if c in res.columns]
            return res[actual_cols].rename(columns=h_map)

        sheets_data = {
            '코스피_상승': process_data('KOSPI', True),
            '코스닥_상승': process_data('KOSDAQ', True),
            '코스피_하락': process_data('KOSPI', False),
            '코스닥_하락': process_data('KOSDAQ', False)
        }

        file_name = f"{target_date_str}_{report_type}_국내리포트.xlsx"
        
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid")
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_white = PatternFill(fill_type=None) 
        font_white = Font(color="FFFFFF", bold=True)
        font_black = Font(color="000000")

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                col_list = list(data.columns)
                name_idx = col_list.index('종목명') + 1
                ratio_idx = col_list.index('전일대비(%)') + 1

                for row in range(2, ws.max_row + 1):
                    val = ws.cell(row=row, column=ratio_idx).value
                    ratio_val = abs(float(val)) if val is not None else 0
                    name_cell = ws.cell(row=row, column=name_idx)

                    # 색상 기준
                    if ratio_val >= 25: 
                        name_cell.fill = fill_red
                        name_cell.font = font_white
                    elif ratio_val >= 20: 
                        name_cell.fill = fill_orange
                        name_cell.font = font_black
                    elif ratio_val >= 10: 
                        name_cell.fill = fill_yellow
                        name_cell.font = font_black
                    else:
                        name_cell.fill = fill_white
                        name_cell.font = font_black

                    # 전체 셀 가운데 정렬 및 숫자 포맷
                    for col in range(1, len(col_list) + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        if isinstance(cell.value, (int, float)):
                            if col_list[col-1] == '전일대비(%)':
                                cell.number_format = '0.00'
                            else:
                                cell.number_format = '#,##0'

                for i in range(1, len(col_list) + 1):
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 16

        bot = Bot(token=TOKEN)
        async with bot:
            count_up = len(sheets_data['코스피_상승']) + len(sheets_data['코스닥_상승'])
            count_down = len(sheets_data['코스피_하락']) + len(sheets_data['코스닥_하락'])
            
            msg = (f"📅 {target_date_str} {report_type} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {count_up}개\n"
                   f"📉 하락(5%↓): {count_down}개\n\n"
                   f"💡 고가/저가 정보가 추가되었습니다.\n"
                   f"⚪ 5%↑ | 🟡 10%↑ | 🟠 20%↑ | 🔴 25%↑")
            
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        
        print(f"--- [성공] {file_name} 전송 완료 ---")

    except Exception as e:
        import traceback
        print(f"에러 발생:\n{traceback.format_exc()}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
