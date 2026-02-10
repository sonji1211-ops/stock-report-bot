import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill

# [설정] 직접 입력 모드 (가장 확실함)
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    # 1. 한국 시간(KST)으로 기준 설정
    # 깃허브 서버(UTC)보다 9시간 빠른 한국 시간으로 날짜를 잡습니다.
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')
    
    # 보고서 타입 (오늘 요일에 따라 자동 결정)
    report_type = "주간" if now.weekday() == 5 else "일일"

    try:
        print(f"--- 분석 시작: {target_date_str} (한국시간 기준) ---")
        
        # 2. 데이터 수집 (KRX 전체 종목)
        df = fdr.StockListing('KRX')
        if df is None or df.empty:
            print("데이터를 가져오는 데 실패했습니다.")
            return

        # 3. 필수 컬럼 정리 및 숫자 변환
        cols = df.columns.tolist()
        chg_amt_col = next((c for c in ['Change', 'Changes', 'ChgAmt'] if c in cols), None)
        cap_col = next((c for c in ['Marcap', 'Amount', 'MarketCap'] if c in cols), cols[-1])

        needed_cols = ['Open', 'Close', 'Volume', cap_col]
        if chg_amt_col: needed_cols.append(chg_amt_col)
        for c in needed_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        # 4. 등락률 계산
        if chg_amt_col:
            def calculate_ratio(row):
                prev_close = row['Close'] - row[chg_amt_col]
                return (row[chg_amt_col] / prev_close * 100) if prev_close != 0 else 0
            df['Calculated_Ratio'] = df.apply(calculate_ratio, axis=1)
        else:
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), cols[-1])
            df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
            if df['Calculated_Ratio'].max() > 100: df['Calculated_Ratio'] /= 100

        # 5. 한글 매핑 및 필터링 (상승/하락 5% 기준)
        h_map = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Open': '시가', 'Close': '종가(현재가)', 
            'Calculated_Ratio': '전일대비(%)', 'Volume': '거래량'
        }

        def process_data(market, is_up):
            m_df = df[(df['Market'].str.contains(market, na=False)) & (df['Volume'] > 0)].copy()
            if is_up:
                res = m_df[m_df['Calculated_Ratio'] >= 5].copy()
            else:
                res = m_df[m_df['Calculated_Ratio'] <= -5].copy()
            
            res = res.sort_values(by=cap_col, ascending=False)
            actual_cols = [c for c in h_map.keys() if c in res.columns]
            return res[actual_cols].rename(columns=h_map)

        sheets_data = {
            '코스피_상승': process_data('KOSPI', True),
            '코스닥_상승': process_data('KOSDAQ', True),
            '코스피_하락': process_data('KOSPI', False),
            '코스닥_하락': process_data('KOSDAQ', False)
        }

        # 6. 엑셀 파일 생성 및 꾸미기
        file_name = f"{target_date_str}_{report_type}_리포트.xlsx"
        
        fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid")
        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

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

                    # 등락률별 색상 지정
                    if ratio_val >= 30: name_cell.fill = fill_red
                    elif ratio_val >= 20: name_cell.fill = fill_orange
                    elif ratio_val >= 10: name_cell.fill = fill_yellow

                    for col in range(1, len(col_list) + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        
                        # 천단위 콤마 서식
                        if isinstance(cell.value, (int, float)):
                            header_name = col_list[col-1]
                            if header_name in ['시가', '종가(현재가)', '거래량']:
                                cell.number_format = '#,##0'
                            elif header_name == '전일대비(%)':
                                cell.number_format = '0.00'

                for i in range(1, len(col_list) + 1):
                    ws.column_dimensions[chr(64+i)].width = 20

        # 7. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"📅 {target_date_str} 리포트 배달완료!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀에서 종목명 색깔을 확인하세요!\n"
                   f"(🟡10%↑, 🟠20%↑, 🔴30%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        print(f"--- [성공] {file_name} 전송 완료 ---")

    except Exception as e:
        import traceback
        print(f"오류 발생:\n{traceback.format_exc()}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
