import os
import pandas as pd
import requests
import re
import io
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
import time

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_total_market_scan_github():
    """깃허브 환경 최적화: 2,500개 전 종목을 누락 없이 초고속 스캔"""
    # 전 종목 필드 강제 활성화 (시가, 고가, 저가, 등락률 포함)
    base_params = "field=quant&field=open&field=high&field=low&field=frate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/sise/sise_market_sum.naver'
    }
    all_stocks = []
    
    for m_code in [0, 1]: # 코스피, 코스닥
        market_label = "KOSPI" if m_code == 0 else "KOSDAQ"
        # 1. 마지막 페이지 확인 (전수조사 필수 단계)
        res = requests.get(f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}", headers=headers)
        last_page_match = re.findall(r'page=(\d+)', res.text)
        last_page = int(max(map(int, last_page_match))) if last_page_match else 1
        
        # 2. 전 페이지 스캔 (목록에서 직접 추출하여 속도 극대화)
        for page in range(1, last_page + 1):
            try:
                url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}&{base_params}&page={page}"
                resp = requests.get(url, headers=headers, timeout=10)
                dfs = pd.read_html(io.StringIO(resp.text))
                df = dfs[1].dropna(subset=['종목명']).copy()
                
                # 데이터 정제 (숫자 변환 및 기호 제거)
                for col in ['등락률', '현재가', '시가', '고가', '저가', '거래량']:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace('%','').str.replace(',','').str.replace('+','').replace('nan', '0')
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

                for _, row in df.iterrows():
                    all_stocks.append({
                        'Name': str(row['종목명']),
                        'Open': int(row['시가']),
                        'Close': int(row['현재가']),
                        'Low': int(row['저가']),
                        'High': int(row['고가']),
                        'Ratio': float(row['등락률']),
                        'Volume': int(row['거래량']),
                        'Market': market_label
                    })
            except: continue
            time.sleep(0.05) # 깃허브 IP 차단 방지용 미세 지연
            
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    # 깃허브 서버는 UTC 기준이므로 한국 시간(KST)으로 보정
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        print("📡 [GitHub] 전 종목 전수조사 및 디자인 리포트 생성 중...")
        df_final = await get_total_market_scan_github()
        if df_final.empty: return

        report_type = "주간평균" if day_of_week == 6 else "일일"
        if day_of_week == 5: report_type = "일일(금요마감)"

        # 필터링 로직 (거래량 0 제외 및 상하 5% 기준)
        def get_sub(m_name, is_up):
            temp = df_final[(df_final['Market'] == m_name) & (df_final['Volume'] > 0)].copy()
            if is_up:
                return temp[temp['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
            else:
                return temp[temp['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

        sheets = {
            '코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True),
            '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)
        }

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        h_map = {'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        # --- 엑셀 디자인 서식 (지수님 로컬용과 동일하게 맞춤) ---
        header_fill = PatternFill("solid", fgColor="444444")
        header_font = Font(color="FFFFFF", bold=True)
        red, orange, yellow = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
        white_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.drop(columns=['Market']).rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                # 본문 스타일 및 색상 강조
                for row in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(row, 6).value or 0))
                    name_cell = ws.cell(row, 1) # 종목명 셀
                    
                    if val >= 28:
                        name_cell.fill, name_cell.font = red, white_font
                    elif val >= 20:
                        name_cell.fill = orange
                    elif val >= 10:
                        name_cell.fill = yellow
                    
                    for col in range(1, 8):
                        ws.cell(row, col).alignment = Alignment(horizontal='center')
                        ws.cell(row, col).border = border
                        if col in [2, 3, 4, 5, 7]: ws.cell(row, col).number_format = '#,##0'
                        if col == 6: ws.cell(row, col).number_format = '0.00'
                
                # 컬럼 너비 조절
                ws.column_dimensions['A'].width = 20
                for char in "BCDEFG": ws.column_dimensions[char].width = 13

        # 텔레그램 전송
        up_total = len(sheets['코스피_상승']) + len(sheets['코스닥_상승'])
        down_total = len(sheets['코스피_하락']) + len(sheets['코스닥_하락'])
        
        msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
               f"📡 전수조사: 국장 전 종목 ({len(df_final)}개) 완료\n\n"
               f"📈 상승(5%↑): {up_total}개\n"
               f"📉 하락(5%↓): {down_total}개\n\n"
               f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
        
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        
        os.remove(file_name) # 임시 파일 삭제
        print("✅ 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
