import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill

# [보안 설정] 깃허브 설정값에서 정보를 가져옵니다
TOKEN = os.environ.get('TELEGRAM_TOKEN')
CHAT_ID = os.environ.get('TELEGRAM_CHAT_ID')

async def send_smart_report():
    # 한국 시간 설정을 위해 현재 시간에서 9시간을 더합니다 (선택 사항)
    now = datetime.now()
    
    # [테스트를 위해 휴무 체크 주석 처리]
    # if now.weekday() == 6: return # 일요일 휴무

    # 기준일 설정 (월요일이면 금요일 데이터, 그 외엔 전일 데이터)
    target_date = now - timedelta(days=3 if now.weekday() == 0 else 1)
    target_date_str = target_date.strftime('%Y-%m-%d')
    report_type = "주간" if now.weekday() == 5 else "일일"

    try:
        print(f"--- 데이터 수집 시작: {target_date_str} ---")
        # 1. 데이터 수집
        df = fdr.StockListing('KRX')
        if df is None or df.empty: 
            print("데이터가 없습니다.")
            return

        # 2. 컬럼 이름 찾기 (FinanceDataReader 버전에 대응)
        cols = df.columns.tolist()
        chg_amt_col = next((c for c in ['Change', 'Changes', 'ChgAmt'] if c in cols), None)
        cap_col = next((c for c in ['Marcap', 'Amount', 'MarketCap'] if c in cols), cols[-1])

        # 3. 데이터 숫자 변환
        needed_cols = ['Open', 'Close', 'Volume', cap_col]
        if chg_amt_col: needed_cols.append(chg_amt_col)
        for c in needed_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)
        
        # 4. 전일 대비 등락률 계산
        if chg_amt_col:
            def calculate_ratio(row):
                prev_close = row['Close'] - row[chg_amt_col]
                return (row[chg_amt_col] / prev_close * 100) if prev_close != 0 else 0
            df['Calculated_Ratio'] = df.apply(calculate_ratio, axis=1)
        else:
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), cols[-1])
            df['Calculated_Ratio'] = pd.to_numeric(df[ratio_col], errors='coerce').fillna(0)
            if df['Calculated_Ratio'].max() > 100: df['Calculated_Ratio'] /= 100

        # 5. 한글 매핑 및 필터링
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

        # 6. 엑셀 저장 및 스타일 적용
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
                    ratio_val = abs(ws.cell(row=row, column=ratio_idx).value or 0)
                    name_cell = ws.cell(row=row, column=name_idx)

                    if ratio_val >= 30: name_cell.fill = fill_red
                    elif ratio_val >= 20: name_cell.fill = fill_orange
                    elif ratio_val >= 10: name_cell.fill = fill_yellow

                    for col in range(1, len(col_list) + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.alignment = Alignment(horizontal='center', vertical='center')

                for i in range(1, len(col_list) + 1):
                    ws.column_dimensions[chr(64+i)].width = 20

        # 7. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            msg = (f"📅 {target_date_str} 리포트\n"
                   f"🚀 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])} / 📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)
        print(f"--- [성공] 전송 완료 ---")

    except Exception as e:
        print(f"오류 상세: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
