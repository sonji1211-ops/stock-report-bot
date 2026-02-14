import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font
import traceback

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_smart_report():
    bot = Bot(token=TOKEN)
    # 한국 시간 설정 (UTC+9)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday() 
    
    # 1. 보고서 날짜 설정 (토요일 실행 시 어제인 금요일 데이터 조준)
    if day_of_week == 6: # 일요일 실행
        report_type = "주간누적(월~금)"
        end_date = (now - timedelta(days=2)).strftime('%Y-%m-%d')   # 금요일
        start_date = (now - timedelta(days=6)).strftime('%Y-%m-%d') # 월요일
    elif day_of_week == 5: # 토요일 실행 (어제 금요일 데이터)
        report_type = "일일(금요일마감)"
        end_date = (now - timedelta(days=1)).strftime('%Y-%m-%d')
        start_date = end_date
    else: # 평일
        report_type = "일일"
        end_date = now.strftime('%Y-%m-%d')
        start_date = end_date

    try:
        print(f"--- {report_type} 리포트 생성 시작 ---")
        
        # 2. 데이터 수집 시도 (재시도 및 이중화)
        df_base = None
        for i in range(3):
            try:
                df_base = fdr.StockListing('KRX')
                if df_base is not None and not df_base.empty:
                    break
            except:
                await asyncio.sleep(3)
        
        # 만약 KRX 전체 목록이 실패하면 주요 대형주 위주로라도 강제 구성
        if df_base is None or df_base.empty:
            print("KRX 서버 응답 없음 - 수동 데이터 수집 모드 전환")
            # 최소한의 데이터라도 보내기 위해 코스피 200 등 주요 리스트 대체 시도
            try:
                df_base = fdr.StockListing('KOSPI') 
            except:
                async with bot: await bot.send_message(CHAT_ID, "❌ 현재 거래소 데이터 서버가 완전히 닫혀 있습니다.")
                return

        # 3. 데이터 계산 (일요일 누적 vs 평일/토요일 일일)
        final_list = []
        # 분석 대상: 거래량 상위 600개 (안정성 최우선)
        df_target = df_base.sort_values(by='Volume', ascending=False).head(600)
        
        for idx, row in df_target.iterrows():
            try:
                # 지정된 날짜 범위의 데이터를 가져옴
                d_hist = fdr.DataReader(row['Code'], start_date, end_date)
                if not d_hist.empty and len(d_hist) >= 1:
                    # 일요일 주간 누적은 시작일과 종료일 비교
                    if day_of_week == 6 and len(d_hist) >= 2:
                        open_p = d_hist.iloc[0]['Open']
                        close_p = d_hist.iloc[-1]['Close']
                    else:
                        # 평일/토요일은 전일 종가 대비 당일 종가 (또는 금요일 데이터)
                        if len(d_hist) >= 2:
                            open_p = d_hist.iloc[-2]['Close']
                            close_p = d_hist.iloc[-1]['Close']
                        else: continue # 데이터 부족 시 패스
                        
                    ratio = round(((close_p - open_p) / open_p) * 100, 2)
                    
                    final_list.append({
                        '종목코드': row['Code'], '종목명': row['Name'], '시장': row['Market'],
                        '시가': d_hist.iloc[-1]['Open'], '고가': d_hist['High'].max(),
                        '저가': d_hist['Low'].min(), '종가': close_p,
                        '등락률(%)': ratio, '거래량': d_hist.iloc[-1]['Volume']
                    })
            except: continue

        df_final = pd.DataFrame(final_list)
        if df_final.empty:
            async with bot: await bot.send_message(CHAT_ID, f"❌ {report_type} 분석 결과 데이터가 비어있습니다.")
            return

        # 4. 엑셀 분류 (지수님 요청 5% 기준)
        def get_subset(is_up, market):
            cond = (df_final['시장'].str.contains(market))
            if is_up:
                return df_final[cond & (df_final['등락률(%)'] >= 5)].sort_values(by='등락률(%)', ascending=False)
            else:
                return df_final[cond & (df_final['등락률(%)'] <= -5)].sort_values(by='등락률(%)', ascending=True)

        sheets = {
            '코스피_상승': get_subset(True, 'KOSPI'), '코스닥_상승': get_subset(True, 'KOSDAQ'),
            '코스피_하락': get_subset(False, 'KOSPI'), '코스닥_하락': get_subset(False, 'KOSDAQ')
        }

        # 5. 엑셀 파일 생성 및 스타일링 (28% 빨간색🔴 포함)
        file_name = f"{now.strftime('%Y-%m-%d')}_국내리포트.xlsx"
        fill_red, fill_orange, fill_yellow = PatternFill(start_color="FF0000", fill_type="solid"), PatternFill(start_color="FFCC00", fill_type="solid"), PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row=row, column=8).value or 0)) # 등락률
                    name_cell = ws.cell(row=row, column=2) # 종목명
                    
                    # 지수님 전용 색상 가이드 (10/20/28)
                    if val >= 28:
                        name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20:
                        name_cell.fill = fill_orange
                    elif val >= 10:
                        name_cell.fill = fill_yellow
                    
                    for c in range(1, 10):
                        cell = ws.cell(row=row, column=c)
                        cell.alignment = Alignment(horizontal='center')
                        if c == 8: cell.number_format = '0.00'
                        elif c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                for i in range(1, 10): ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 15

        # 6. 전송
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 리포트\n\n"
                   f"📈 상승: {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])} / 📉 하락: {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}\n"
                   f"💡 가이드: (🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        print(f"오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
