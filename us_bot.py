import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930" 

async def send_us_smart_report():
    now = datetime.utcnow() + timedelta(hours=9)
    target_date_str = now.strftime('%Y-%m-%d')

    # [수정] 가장 안정적인 데이터 소스 심볼로 재배치
    indices = {
        '나스닥': 'IXIC', 
        'S&P500': 'US500', 
        '필라델피아 반도체': 'SOX'
    }

    try:
        print(f"--- 미국 증시 분석 시작: {target_date_str} ---")
        report_data = []
        summary_text = f"🇺🇸 {target_date_str} 미국 증시 마감\n\n"

        for name, symbol in indices.items():
            try:
                # 야후 파이낸스 에러를 피하기 위해 데이터 로딩 시도
                df = fdr.DataReader(symbol)
                
                # 만약 데이터를 못 가져왔다면 다른 심볼로 재시도
                if df is None or df.empty:
                    alt_symbols = {'나스닥': 'NASDAQ', '필라델피아 반도체': 'PHLX Semiconductor'}
                    if name in alt_symbols:
                        df = fdr.DataReader(alt_symbols[name])
                
                if df is not None and not df.empty:
                    last = df.iloc[-1]
                    prev = df.iloc[-2]
                    
                    close_val = float(last['Close'])
                    change_val = close_val - float(prev['Close'])
                    chg_ratio = (change_val / float(prev['Close'])) * 100
                    
                    icon = "📈" if change_val > 0 else "📉"
                    summary_text += f"{icon} {name}: {chg_ratio:+.2f}%\n"

                    report_data.append({
                        '지수명': name,
                        '현재지수': close_val,
                        '전일대비': change_val,
                        '등락률(%)': chg_ratio,
                        '시가': last['Open'],
                        '고가': last['High'],
                        '저가': last['Low']
                    })
                else:
                    print(f"{name} 데이터 수집 실패")
            except:
                print(f"{name} 수집 중 오류 발생 - 건너뜀")
                continue

        if not report_data:
            print("데이터가 하나도 없습니다.")
            return

        # 3. 엑셀 파일 생성
        file_name = f"{target_date_str}_미국증시_리포트.xlsx"
        df_final = pd.DataFrame(report_data)

        fill_red = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        fill_blue = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        white_font = Font(color="FFFFFF", bold=True)

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='미국지수', index=False)
            ws = writer.sheets['미국지수']
            
            for row in range(2, ws.max_row + 1):
                ratio_val = ws.cell(row=row, column=4).value 
                name_cell = ws.cell(row=row, column=1) 
                if ratio_val:
                    if ratio_val > 0:
                        name_cell.fill = fill_red
                        name_cell.font = white_font
                    elif ratio_val < 0:
                        name_cell.fill = fill_blue
                        name_cell.font = white_font
                
                for col in range(1, 8):
                    cell = ws.cell(row=row, column=col)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00'

            for i in range(1, 8):
                ws.column_dimensions[chr(64+i)].width = 20

        # 4. 텔레그램 전송
        bot = Bot(token=TOKEN)
        async with bot:
            summary_text += "\n📊 상세 내용은 엑셀 파일을 확인하세요!"
            with open(file_name, 'rb') as f:
                await bot.send_document(chat_id=CHAT_ID, document=f, caption=summary_text)
        print(f"--- [성공] 전송 완료 ---")

    except Exception as e:
        print(f"최종 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_us_smart_report())
