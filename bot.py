import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import time

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def fetch_stock_data(row, start_d, end_d, semaphore):
    """서버 차단을 피하기 위해 요청 간격에 미세한 지연을 추가합니다."""
    async with semaphore:
        try:
            # 야후 소스를 사용하되 요청 전 0.5초 대기 (차단 방지)
            await asyncio.sleep(0.5) 
            suffix = ".KS" if row['Market'] == 'KOSPI' else ".KQ"
            ticker = f"{row['Code']}{suffix}"
            
            df = fdr.DataReader(ticker, start_d, end_d)
            if df is None or len(df) < 2:
                return None
            
            last_c = float(df.iloc[-1]['Close'])
            prev_c = float(df.iloc[-2]['Close'])
            ratio = round(((last_c - prev_c) / prev_c) * 100, 2)
            
            return {
                'Code': row['Code'], 'Name': row['Name'], 
                'Open': df.iloc[-1]['Open'], 'Close': last_c,
                'Low': df['Low'].min(), 'High': df['High'].max(), 
                'Ratio': ratio, 'Volume': df.iloc[-1]['Volume'],
                'Market': row['Market']
            }
        except:
            return None

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 

    try:
        # [핵심 수정] StockListing 자체가 막히는 경우를 대비해 예외 처리
        try:
            df_base = fdr.StockListing('KRX')
        except:
            # KRX가 막히면 KOSPI/KOSDAQ 따로 시도
            df_base = pd.concat([fdr.StockListing('KOSPI'), fdr.StockListing('KOSDAQ')])

        if df_base is None or df_base.empty:
            return

        # [안전 장치] 동시 접속을 5개로 대폭 제한 (천천히 가져오기)
        sem = asyncio.Semaphore(5)
        
        if day_of_week == 6: # 일요일
            report_type, analysis_info = "주간평균", "시총 상위 500"
            end_d = (now - timedelta(days=2)).strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=6)).strftime('%Y-%m-%d')
            df_target = df_base.sort_values(by='Marcap', ascending=False).head(500).copy()
        else: # 평일
            report_type, analysis_info = "일일", "전 종목 전수조사"
            if day_of_week == 5: report_type = "일일(금요마감)"
            end_d = now.strftime('%Y-%m-%d')
            start_d = (now - timedelta(days=5)).strftime('%Y-%m-%d')
            # [임시 조치] 만약 전수조사에서 계속 터진다면 head(1000)으로 줄여서 테스트해보세요
            df_target = df_base.copy()

        tasks = [fetch_stock_data(row, start_d, end_d, sem) for _, row in df_target.iterrows()]
        results = await asyncio.gather(*tasks)
        df_final = pd.DataFrame([r for r in results if r is not None])
        
        if df_final.empty:
            print("데이터 수집 실패: 모든 요청이 차단되었습니다.")
            return

        # --- 이하 디자인 및 전송 로직 동일 ---
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_data(m_name, is_up):
            temp = df_final[df_final['Market'].str.contains(m_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {'코스피_상승': get_data('KOSPI', True), '코스닥_상승': get_data('KOSDAQ', True),
                  '코스피_하락': get_data('KOSPI', False), '코스닥_하락': get_data('KOSDAQ', False)}

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        # (디자인 생략 - 위 코드와 동일하게 유지됨)
        # [생략된 디자인 코드는 위에서 드린 '가독성 최적화' 버전과 100% 같습니다]
        # ... 디자인 로직 적용 ...

        # [디자인 적용을 위해 위 코드의 엑셀 디자인 부분을 그대로 넣어주세요]
        header_fill = PatternFill("solid", fgColor="444444")
        header_font = Font(color="FFFFFF", bold=True)
        fills = [PatternFill("solid", fgColor="FF0000"), PatternFill("solid", fgColor="FFBB00"), PatternFill("solid", fgColor="FFFF00")]
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                for row in range(2, ws.max_row + 1):
                    ratio = abs(float(ws.cell(row, 7).value or 0))
                    name_cell = ws.cell(row, 2)
                    if ratio >= 28: name_cell.fill, name_cell.font = fills[0], Font(color="FFFFFF", bold=True)
                    elif ratio >= 20: name_cell.fill = fills[1]
                    elif ratio >= 10: name_cell.fill = fills[2]
                    for col in range(1, 9):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = border
                        if col in [3,4,5,6,8]: ws.cell(row, col).number_format = '#,##0'
                        if col == 7: ws.cell(row, col).number_format = '0.00'
                ws.column_dimensions['B'].width = 20
                for char in "ACDEFGH": ws.column_dimensions[char].width = 12

        async with bot:
            date_str = f"{start_d}~{end_d}" if day_of_week == 6 else now.strftime('%Y-%m-%d')
            msg = (f"📦 *[{report_type}] 리포트 (차단우회 모드)*\n\n"
                   f"📅 *날짜:* {date_str}\n"
                   f"🔍 *대상:* {analysis_info}\n"
                   f"───\n"
                   f"📈 *상승(5%↑):* {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])}개\n"
                   f"📉 *하락(5%↓):* {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}개\n\n"
                   f"💡 *안내:* 서버 IP 차단으로 인해 수집 속도가 제한되었습니다.")
            await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg, parse_mode="Markdown")

    except Exception as e:
        print(f"최종 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
