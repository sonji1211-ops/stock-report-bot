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

# [설정] 텔레그램 정보 (지수님 정보 고정)
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def get_naver_realtime_data(limit=700):
    """국장 차단을 뚫고 실시간 데이터를 수집하는 핵심 함수"""
    list_url = "https://finance.naver.com/sise/sise_market_sum.naver"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/'
    }
    all_stocks = []
    
    # 코스피(0), 코스닥(1) 순회
    for m_code in [0, 1]:
        market_label = "KOSPI" if m_code == 0 else "KOSDAQ"
        # 상위 종목 위주로 페이지당 50개씩 수집
        for page in range(1, (limit // 50) + 2):
            try:
                resp = requests.get(list_url, params={'sosok': m_code, 'page': page}, headers=headers, timeout=10)
                item_codes = re.findall(r'code=(\d{6})', resp.text)
                item_codes = list(dict.fromkeys(item_codes))
                
                dfs = pd.read_html(io.StringIO(resp.text))
                df_list = dfs[1].dropna(subset=['종목명']).copy()
                
                for i, (idx, row) in enumerate(df_list.iterrows()):
                    if i >= len(item_codes): break
                    code = item_codes[i]
                    try:
                        # 네이버 실시간 상세 API (등락률 부호 및 실제 시/고/저/종가)
                        api_url = f"https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:{code}"
                        api_resp = requests.get(api_url, headers=headers, timeout=5).json()
                        item = api_resp['result']['areas'][0]['datas'][0]
                        
                        all_stocks.append({
                            'Code': code, 'Name': row['종목명'], 
                            'Open': int(item['sv']), 'Close': int(item['nv']),
                            'Low': int(item['lv']), 'High': int(item['hv']), 
                            'Ratio': float(item['cr']), 'Volume': int(item['aq']),
                            'Market': market_label
                        })
                    except: continue
            except: continue
            time.sleep(0.1) # 깃허브 서버 차단 방지 미세 지연
            
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    """메인 실행 함수: 데이터 분류 -> 엑셀 생성 -> 텔레그램 전송"""
    bot = Bot(token=TOKEN)
    # 한국 시간(KST) 설정
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        # 요일별 수집 범위 설정 (일요일: 500개 / 평일: 700개)
        limit_count = 500 if day_of_week == 6 else 700
        report_type = "주간평균" if day_of_week == 6 else "일일"
        if day_of_week == 5: report_type = "일일(금요마감)"
        
        print(f"📡 {report_type} 리포트 분석을 시작합니다...")
        df_final = await get_naver_realtime_data(limit=limit_count)

        if df_final.empty: 
            print("❌ 수집된 데이터가 없습니다.")
            return

        # 데이터 필터링 (상승/하락 5% 기준)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        def get_filtered_data(m_name, is_up):
            temp = df_final[df_final['Market'].str.contains(m_name, na=False)].copy()
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {
            '코스피_상승': get_filtered_data('KOSPI', True), '코스닥_상승': get_filtered_data('KOSDAQ', True),
            '코스피_하락': get_filtered_data('KOSPI', False), '코스닥_하락': get_filtered_data('KOSDAQ', False)
        }

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        
        # --- 엑셀 디자인 설정 (지수님 요청 사항) ---
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
                
                # 헤더 스타일링
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                # 본문 스타일링 및 색상 강조
                for row in range(2, ws.max_row + 1):
                    ratio_val = ws.cell(row, 7).value
                    ratio = abs(float(ratio_val if ratio_val is not None else 0))
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
                        if col in [3, 4, 5, 6, 8]: ws.cell(row, col).number_format = '#,##0'
                        if col == 7: ws.cell(row, col).number_format = '0.00'
                
                ws.column_dimensions['B'].width = 20
                for char in "ACDEFGH": ws.column_dimensions[char].width = 13

        # --- 텔레그램 전송 ---
        msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
               f"📡 분석: {len(df_final)}개 종목 수집 완료\n\n"
               f"📈 상승(5%↑): {len(sheets['코스피_상승'])+len(sheets['코스닥_상승'])}개\n"
               f"📉 하락(5%↓): {len(sheets['코스피_하락'])+len(sheets['코스닥_하락'])}개\n\n"
               f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
        
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        
        os.remove(file_name) # 서버 임시 파일 삭제
        print(f"✅ {file_name} 전송 완료!")

    except Exception as e:
        print(f"❌ 오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
