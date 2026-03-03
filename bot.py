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

async def fetch_stock_data(row, start_d, end_d, semaphore):
    """증권사 및 포털 소스를 다각화하여 초저속으로 수집합니다."""
    async with semaphore:
        code, name, market = row['Code'], row['Name'], row['Market']
        
        # ⚠️ 차단 방지: 증권사 서버가 의심하지 않도록 1.2~2.8초 랜덤 휴식
        await asyncio.sleep(random.uniform(1.2, 2.8))
        
        df = None
        # 데이터 소스 우선순위: daum(증권사 연동) -> yahoo -> naver
        for src in ['daum', 'yahoo', 'naver']:
            try:
                ticker = code
                if src == 'yahoo':
                    ticker = f"{code}.KS" if market == 'KOSPI' else f"{code}.KQ"
                
                df = fdr.DataReader(ticker, start_d, end_d)
                if df is not None and len(df) >= 2:
                    break 
            except:
                continue
        
        if df is None or len(df) < 2:
            return None

        try:
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
        except:
            return None

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        # 1. 실시간 종목 리스트 확보 (실패 시 우회로 가동)
        try:
            df_base = fdr.StockListing('KRX')
        except:
            df_base = pd.concat([fdr.StockListing('KOSPI'), fdr.StockListing('KOSDAQ')])

        if df_base is None or df_base.empty:
            return

        # 2. 요일별 분석 대상 (지수님 원본 로직 100%)
        if day_of_week == 6: # 일요일: 주간 시총 상위 500
            report_type, analysis_info = "주간평균", "시총 상위 500"
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
        else: # 평일: 전 종목 전수조사
            report_type, analysis_info = "일일", "전 종목 전수조사"
            if day_of_week == 5: report_type = "일일(금요마감)"
            df_target = df_base.copy()
            end_d = now.strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=4)).strftime('%Y-%m-%d')

        # 3. 비동기 수집 (안전하게 동시 2개씩만 찔러서 차단 회피)
        sem = asyncio.Semaphore(2)
        print(f"[{report_type}] 증권사급 우회 모드 가동... 총 {len(df_target)}개 분석")
        
        tasks = [fetch_stock_data(row, start_d, end_d, sem) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        df_final = pd.DataFrame([r for r in results if r is not None])

        if df_final.empty:
            print("데이터를 하나도 가져오지 못했습니다. 환경을 점검해 주세요.")
            return

        # 4. 분류 및 필터링 (5% 기준)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_data(m_name, is_up):
            temp = df_final[df_final['Market'].str.contains(m_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {'코스피_상승': get_data('KOSPI', True), '코스닥_상승': get_data('KOSDAQ', True),
                  '코스피_하락': get_data('KOSPI', False), '코스닥_하락': get_data('KOSDAQ', False)}

        # 5. 엑셀 디자인 (누락 없이 디테일하게 복구)
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
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
                
                # 상단 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                # 본문 스타일 및 조건부 색상
                for row in range(2, ws.max_row + 1):
                    ratio = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    
                    if ratio >= 28:
                        name_cell.fill, name_cell.font = fills[0], Font(color="FFFFFF", bold=True)
                    elif ratio >= 20:
                        name_cell.fill = fills[1]
                    elif ratio >= 10:
                        name_cell.fill = fills[2]
                    
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = border
                        if col in [3,4,5,6,8]: ws.cell(row, col).number_format = '#,##0'
                        if col == 7: ws.cell(row, col).number_format = '0.00'
                
                ws.column_dimensions['B'].width = 20
                for char in "ACDEFGH": ws.column_dimensions[char].width = 13

        # 6. 텔레그램 전송
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} *[{report_type}] 리포트*\n"
                   f"📡 데이터 소스: 증권사/포털 통합 우회\n\n"
                   f"📈 상승(5%↑): {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}개\n\n"
                   f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg, parse_mode="Markdown")

    except Exception as e:
        print(f"최종 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
