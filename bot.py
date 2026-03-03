import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import random

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def fetch_stock_safe(row, start_d, end_d, semaphore):
    async with semaphore:
        code, name, market = row['Code'], row['Name'], row['Market']
        # ⚠️ 차단 방지를 위해 하나씩 천천히 (랜덤 지연 1.5~3초)
        await asyncio.sleep(random.uniform(1.5, 3.0))
        
        for src in ['yahoo', 'google']:
            try:
                ticker = f"{code}.KS" if market == 'KOSPI' else f"{code}.KQ"
                if src == 'google': ticker = f"KRX:{code}"
                
                df = fdr.DataReader(ticker, start_d, end_d)
                if df is not None and len(df) >= 2:
                    last_c = float(df.iloc[-1]['Close'])
                    prev_c = float(df.iloc[-2]['Close'])
                    ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
                    return {
                        'Code': code, 'Name': name, 
                        'Open': df.iloc[-1]['Open'], 'Close': last_c,
                        'Low': df['Low'].min(), 'High': df['High'].max(), 
                        'Ratio': ratio, 'Volume': df.iloc[-1]['Volume'],
                        'Market': market
                    }
            except: continue
        return None

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        # 1. 종목 리스트 확보
        df_base = fdr.StockListing('KRX')
        
        # 2. 요일별 범위 설정 (GitHub 생존을 위해 일일 700개로 최적화)
        if day_of_week == 6: # 일요일: 주간 시총 상위 500
            report_type, analysis_info = "주간평균", "시총 상위 500"
            df_target = df_base.sort_values('Marcap', ascending=False).head(500).copy()
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
        else: # 평일: 일일 시총 상위 700 (차단 방지 마지노선)
            report_type, analysis_info = "일일", "상위 700 종목"
            if day_of_week == 5: report_type = "일일(금요마감)"
            df_target = df_base.sort_values('Marcap', ascending=False).head(700).copy()
            start_d = (now - timedelta(days=4)).strftime('%Y-%m-%d')
        
        end_d = now.strftime('%Y-%m-%d')

        # 3. 비동기 수집 (안전하게 1개씩 순차 처리)
        sem = asyncio.Semaphore(1) 
        tasks = [fetch_stock_safe(row, start_d, end_d, sem) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        df_final = pd.DataFrame([r for r in results if r is not None])

        if df_final.empty: return

        # 4. 분류 및 필터링 (5% 기준)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_data(m_name, is_up):
            temp = df_final[df_final['Market'].str.contains(m_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {'코스피_상승': get_data('KOSPI', True), '코스닥_상승': get_data('KOSDAQ', True),
                  '코스피_하락': get_data('KOSPI', False), '코스닥_하락': get_data('KOSDAQ', False)}

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        
        # 5. 엑셀 디자인 (누락 없이 완벽 복구)
        header_fill = PatternFill("solid", fgColor="444444")
        header_font = Font(color="FFFFFF", bold=True)
        fills = [
            PatternFill("solid", fgColor="FF0000"), # 28%↑ 빨강
            PatternFill("solid", fgColor="FFBB00"), # 20%↑ 주황
            PatternFill("solid", fgColor="FFFF00")  # 10%↑ 노랑
        ]
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                # 본문 스타일 (정렬, 테두리, 색상, 포맷)
                for row in range(2, ws.max_row + 1):
                    ratio = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    
                    # 등락률 강조 색상
                    if ratio >= 28:
                        name_cell.fill, name_cell.font = fills[0], Font(color="FFFFFF", bold=True)
                    elif ratio >= 20:
                        name_cell.fill = fills[1]
                    elif ratio >= 10:
                        name_cell.fill = fills[2]
                    
                    # 모든 셀 정렬 및 테두리 적용
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = border
                        # 천 단위 콤마
                        if col in [3, 4, 5, 6, 8]:
                            ws.cell(row, col).number_format = '#,##0'
                        # 소수점 2자리
                        if col == 7:
                            ws.cell(row, col).number_format = '0.00'
                
                # 열 너비 조절
                ws.column_dimensions['B'].width = 20
                for char in "ACDEFGH":
                    ws.column_dimensions[char].width = 13

        # 6. 전송
        async with bot:
            msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
                   f"📡 분석: {len(df_final)}개 종목 성공\n\n"
                   f"📈 상승(5%↑): {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}개\n\n"
                   f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg, parse_mode="Markdown")

    except Exception as e:
        print(f"최종 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
