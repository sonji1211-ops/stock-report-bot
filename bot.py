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
import random

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_total_market_scan_github():
    """요구사항 1: 전수조사 (모든 페이지 스캔으로 하락주 누락 방지)"""
    base_params = "field=quant&field=open&field=high&field=low&field=frate"
    user_agents = [
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/122.0.0.0 Safari/537.36',
        'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Chrome/121.0.0.0 Safari/537.36'
    ]
    all_stocks = []
    
    for m_code in [0, 1]: # 0: 코스피, 1: 코스닥
        market_label = "KOSPI" if m_code == 0 else "KOSDAQ"
        try:
            headers = {'User-Agent': random.choice(user_agents), 'Referer': 'https://finance.naver.com/'}
            res = requests.get(f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}", headers=headers, timeout=10)
            last_page_match = re.findall(r'page=(\d+)', res.text)
            last_page = int(max(map(int, last_page_match))) if last_page_match else 1
            
            print(f"📡 {market_label} 전수조사 시작 ({last_page}페이지)...")

            for page in range(1, last_page + 1):
                url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}&{base_params}&page={page}"
                success = False
                for attempt in range(3): # 차단 시 3번 재시도
                    try:
                        headers = {'User-Agent': random.choice(user_agents), 'Referer': 'https://finance.naver.com/sise/sise_market_sum.naver'}
                        resp = requests.get(url, headers=headers, timeout=10)
                        if "종목명" in resp.text:
                            dfs = pd.read_html(io.StringIO(resp.text))
                            if len(dfs) >= 2:
                                df = dfs[1].dropna(subset=['종목명']).copy()
                                success = True
                                break
                        time.sleep(0.5)
                    except: time.sleep(1)
                
                if success:
                    for col in ['등락률', '현재가', '시가', '고가', '저가', '거래량']:
                        if col in df.columns:
                            df[col] = df[col].astype(str).str.replace('%','').str.replace(',','').str.replace('+','').replace('nan', '0')
                            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

                    for _, row in df.iterrows():
                        all_stocks.append({
                            'Name': str(row['종목명']), 'Open': int(row['시가']), 'Close': int(row['현재가']),
                            'Low': int(row['저가']), 'High': int(row['고가']), 'Ratio': float(row['등락률']),
                            'Volume': int(row['거래량']), 'Market': market_label
                        })
                else: print(f"⚠️ {market_label} {page}페이지 수집 실패")
        except Exception as e: print(f"❌ {market_label} 오류: {e}")
            
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    
    try:
        df_final = await get_total_market_scan_github()
        if df_final.empty: return

        report_type = "주간평균" if now.weekday() == 6 else "일일"
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"

        # 필터링 로직 (상승/하락 5% 기준)
        def get_sub(m_name, is_up):
            temp = df_final[(df_final['Market'] == m_name) & (df_final['Volume'] > 0)].copy()
            cond = (temp['Ratio'] >= 5.0) if is_up else (temp['Ratio'] <= -5.0)
            return temp[cond].sort_values('Ratio', ascending=not is_up)

        sheets = {
            '코스피_상승': get_sub('KOSPI', True), '코스닥_상승': get_sub('KOSDAQ', True),
            '코스피_하락': get_sub('KOSPI', False), '코스닥_하락': get_sub('KOSDAQ', False)
        }

        # 요구사항 2: 디자인 디테일 (색상, 테두리, 콤마)
        h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
        red, orange, yellow = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
        header_fill = PatternFill("solid", "444444")
        white_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.drop(columns=['Market']).rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 헤더 디자인
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, white_font, Alignment(horizontal='center')
                
                # 요구사항 3: 본문 서식 (콤마, 소수점, 색상)
                for r in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(r, 6).value or 0))
                    # 종목명 색상 강조 규칙
                    if val >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_font
                    elif val >= 20: ws.cell(r, 1).fill = orange
                    elif val >= 10: ws.cell(r, 1).fill = yellow
                    
                    for c in range(1, 8):
                        ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                        # 천 단위 콤마 적용
                        if c in [2, 3, 4, 5, 7]: ws.cell(r, c).number_format = '#,##0'
                        # 등락률 소수점 2자리 적용
                        if c == 6: ws.cell(r, c).number_format = '0.00'
                
                ws.column_dimensions['A'].width = 18
                for char in "BCDEFG": ws.column_dimensions[char].width = 12

        # 텔레그램 메시지
        up_cnt = len(sheets['코스피_상승']) + len(sheets['코스닥_상승'])
        down_cnt = len(sheets['코스피_하락']) + len(sheets['코스닥_하락'])
        msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
               f"📊 분석: 전 종목 ({len(df_final)}개) 전수조사 완료\n"
               f"📈 상승: {up_cnt} / 📉 하락: {down_cnt}\n\n"
               f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑ (모든 하락주 포함)")
        
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        os.remove(file_name)
    except Exception as e: print(f"❌ 오류: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
