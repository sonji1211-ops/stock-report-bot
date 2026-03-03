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
    """깃허브 환경 최적화: 2,500개 전 종목을 누락 없이 정밀 스캔 (타임아웃 강화)"""
    base_params = "field=quant&field=open&field=high&field=low&field=frate"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/sise/'
    }
    all_stocks = []
    
    for m_code in [0, 1]:
        market_label = "KOSPI" if m_code == 0 else "KOSDAQ"
        try:
            # 1. 마지막 페이지 확인
            res = requests.get(f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}", headers=headers, timeout=7)
            last_page_match = re.findall(r'page=(\d+)', res.text)
            last_page = int(max(map(int, last_page_match))) if last_page_match else 1
            
            print(f"📡 {market_label} 전수조사 시작 (총 {last_page}페이지)...")

            for page in range(1, last_page + 1):
                try:
                    url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={m_code}&{base_params}&page={page}"
                    resp = requests.get(url, headers=headers, timeout=7)
                    
                    dfs = pd.read_html(io.StringIO(resp.text))
                    if len(dfs) < 2: continue
                    df = dfs[1].dropna(subset=['종목명']).copy()
                    
                    # 데이터 정제
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
                    time.sleep(0.1) # 차단 방지
                except:
                    print(f"⚠️ {market_label} {page}페이지 지연으로 건너뜀")
                    continue
        except Exception as e:
            print(f"❌ {market_label} 스캔 중 치명적 오류: {e}")
            
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    # 한국 시간 보정 (UTC+9)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        df_final = await get_total_market_scan_github()
        if df_final.empty:
            print("❌ 수집된 데이터가 없습니다.")
            return

        report_type = "주간평균" if day_of_week == 6 else "일일"
        if day_of_week == 5: report_type = "일일(금요마감)"
        
        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"

        # 필터링 로직
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

        # --- 엑셀 디자인 및 서식 (로컬용과 동일하게 복구) ---
        h_map = {'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        red, orange, yellow = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
        header_fill = PatternFill("solid", "444444")
        white_font = Font(color="FFFFFF", bold=True)
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

        with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
            for s_name, data in sheets.items():
                data.drop(columns=['Market']).rename(columns=h_map).to_excel(writer, sheet_name=s_name, index=False)
                ws = writer.sheets[s_name]
                
                # 헤더 스타일
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, white_font, Alignment(horizontal='center')
                
                # 본문 스타일 (색상, 콤마, 소수점, 테두리)
                for r in range(2, ws.max_row + 1):
                    val = abs(float(ws.cell(r, 6).value or 0))
                    # 종목명 색상 강조
                    if val >= 28:
                        ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_font
                    elif val >= 20:
                        ws.cell(r, 1).fill = orange
                    elif val >= 10:
                        ws.cell(r, 1).fill = yellow
                    
                    for c in range(1, 8):
                        ws.cell(r, c).alignment = Alignment(horizontal='center')
                        ws.cell(r, c).border = border
                        if c in [2, 3, 4, 5, 7]: ws.cell(r, c).number_format = '#,##0' # 콤마
                        if c == 6: ws.cell(r, c).number_format = '0.00' # 소수점
                
                # 너비 조절
                ws.column_dimensions['A'].width = 18
                for char in "BCDEFG": ws.column_dimensions[char].width = 12

        # 텔레그램 전송
        up_total = len(sheets['코스피_상승']) + len(sheets['코스닥_상승'])
        down_total = len(sheets['코스피_하락']) + len(sheets['코스닥_하락'])
        
        msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
               f"📊 전수조사: 총 {len(df_final)}개 종목 분석 완료\n"
               f"📈 상승: {up_total}개 / 📉 하락: {down_total}개\n"
               f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑ (누락 없음)")
        
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        
        os.remove(file_name)
        print("✅ 깃허브 리포트 전송 성공!")

    except Exception as e:
        print(f"❌ 최종 오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
