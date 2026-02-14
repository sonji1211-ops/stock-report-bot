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
    day_of_week = now.weekday() # 0:월, 1:화, 2:수, 3:목, 4:금, 5:토, 6:일
    
    # 1. 보고서 타입 및 날짜 범위 설정
    if day_of_week == 6: # 일요일 실행 (월~금 누적 분석)
        report_type = "주간누적(월~금)"
        end_date_str = (now - timedelta(days=2)).strftime('%Y-%m-%d')   # 이번주 금요일
        start_date_str = (now - timedelta(days=6)).strftime('%Y-%m-%d') # 이번주 월요일
    elif day_of_week == 5: # 토요일 실행 (금요일 하루치 마감)
        report_type = "일일(금요일마감)"
        end_date_str = (now - timedelta(days=1)).strftime('%Y-%m-%d')
        start_date_str = end_date_str
    else: # 평일 실행
        report_type = "일일"
        end_date_str = now.strftime('%Y-%m-%d')
        start_date_str = end_date_str

    try:
        print(f"--- {report_type} 분석 시작 ---")
        
        # 2. KRX 종목 리스트 수집 (최대 5회 재시도)
        df_base = None
        for i in range(5):
            try:
                df_base = fdr.StockListing('KRX')
                if df_base is not None and not df_base.empty:
                    break
            except Exception as e:
                print(f"데이터 수집 재시도 중... ({i+1}/5) 에러: {e}")
                await asyncio.sleep(5)
        
        if df_base is None or df_base.empty:
            async with bot:
                await bot.send_message(CHAT_ID, "❌ [국장] 현재 KRX 서버에서 데이터를 불러올 수 없습니다. 잠시 후 Actions를 다시 실행해주세요.")
            return

        # 3. 데이터 계산 로직
        if day_of_week == 6: # [일요일 전용] 주간 누적 수익률 계산
            weekly_data = []
            # 안정성을 위해 거래량 상위 700개 종목 분석
            df_target = df_base.sort_values(by='Volume', ascending=False).head(700)
            for idx, row in df_target.iterrows():
                try:
                    d_hist = fdr.DataReader(row['Code'], start_date_str, end_date_str)
                    if not d_hist.empty and len(d_hist) >= 2:
                        open_p = d_hist.iloc[0]['Open']   # 월요일 시가
                        close_p = d_hist.iloc[-1]['Close'] # 금요일 종가
                        ratio = round(((close_p - open_p) / open_p) * 100, 2)
                        
                        weekly_data.append({
                            'Code': row['Code'], 'Name': row['Name'], 'Market': row['Market'],
                            'Open': open_p, 'High': d_hist['High'].max(), 
                            'Low': d_hist['Low'].min(), 'Close': close_p,
                            'Calculated_Ratio': ratio, 
                            'Volume': d_hist['Volume'].mean()
                        })
                except: continue
            df = pd.DataFrame(weekly_data)
        else: # [평일/토요일 전용] 일일 등락률 계산
            cols = df_base.columns.tolist()
            ratio_col = next((c for c in ['ChgPct', 'ChangesRatio', 'FlucRate'] if c in cols), None)
            df_base['Calculated_Ratio'] = pd.to_numeric(df_base[ratio_col], errors='coerce').fillna(0)
            
            # 소수점 단위 보정 (0.03 -> 3.00)
            if df_base['Calculated_Ratio'].abs().max() < 2: 
                df_base['Calculated_Ratio'] *= 100
            
            df = df_base.copy()
            df['Calculated_Ratio'] = df['Calculated_Ratio'].round(2)

        if df.empty:
            async with bot:
                await bot.send_message(CHAT_ID, f"❌ {report_type} 분석할 데이터가 없습니다.")
            return

        # 4. 엑셀 구조 잡기
        h_map = {
            'Code': '종목코드', 'Name': '종목명', 'Market': '시장',
            'Open': '시가', 'High': '고가', 'Low': '저가', 
            'Close': '종가', 'Calculated_Ratio': '등락률(%)', 'Volume': '거래량'
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

        # 5. 엑셀 파일 생성 및 스타일링
        file_name = f"{now.strftime('%Y-%m-%d')}_국내주식_{report_type}.xlsx"
        fill_red = PatternFill(start_color="FF0000", fill_type="solid")
        fill_orange = PatternFill(start_color="FFCC00", fill_type="solid")
        fill_yellow = PatternFill(start_color="FFFF00", fill_type="solid")
        font_white = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets_data.items():
                data.to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                for row in range(2, ws.max_row + 1):
                    # 8번째 열(등락률) 확인
                    val = abs(float(ws.cell(row=row, column=8).value or 0))
                    name_cell = ws.cell(row=row, column=2)
                    
                    # 지수님 커스텀 색상 기준 (10/20/28)
                    if val >= 28:
                        name_cell.fill, name_cell.font = fill_red, font_white
                    elif val >= 20:
                        name_cell.fill = fill_orange
                    elif val >= 10:
                        name_cell.fill = fill_yellow
                    
                    # 셀 정렬 및 숫자 포맷
                    for c in range(1, 10):
                        cell = ws.cell(row=row, column=c)
                        cell.alignment = Alignment(horizontal='center')
                        if c == 8: # 등락률
                            cell.number_format = '0.00'
                        elif c in [4, 5, 6, 7, 9]: # 금액/거래량
                            cell.number_format = '#,##0'
                
                # 열 너비 조정
                for i in range(1, 10):
                    ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = 15

        # 6. 전송
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 국장 리포트 배달!\n\n"
                   f"📈 상승(5%↑): {len(sheets_data['코스피_상승'])+len(sheets_data['코스닥_상승'])}개\n"
                   f"📉 하락(5%↓): {len(sheets_data['코스피_하락'])+len(sheets_data['코스닥_하락'])}개\n\n"
                   f"💡 엑셀 종목명 색상 가이드\n"
                   f"(🟡10%↑, 🟠20%↑, 🔴28%↑)")
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=msg)

    except Exception as e:
        err_msg = traceback.format_exc()
        print(err_msg)
        async with bot:
            await bot.send_message(CHAT_ID, f"⚠️ 국장 분석 중 오류 발생:\n{str(e)}\n\n내용: {err_msg[:150]}...")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
