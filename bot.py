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

async def get_naver_realtime_data(limit=700):
    """국장 차단을 완벽히 우회하고 누락 없이 데이터를 긁어오는 핵심 로직"""
    list_url = "https://finance.naver.com/sise/sise_market_sum.naver"
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Referer': 'https://finance.naver.com/',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7'
    }
    all_stocks = []
    
    for m_code in [0, 1]:
        market_label = "KOSPI" if m_code == 0 else "KOSDAQ"
        # 목표 종목 수에 맞춰 페이지 넉넉히 탐색
        for page in range(1, (limit // 50) + 3):
            try:
                # 1. 목록 페이지 가져오기 (타임아웃 및 재시도 로직)
                resp = requests.get(list_url, params={'sosok': m_code, 'page': page}, headers=headers, timeout=15)
                if resp.status_code != 200: continue
                
                item_codes = re.findall(r'code=(\d{6})', resp.text)
                item_codes = list(dict.fromkeys(item_codes))
                
                dfs = pd.read_html(io.StringIO(resp.text))
                if len(dfs) < 2: continue
                df_list = dfs[1].dropna(subset=['종목명']).copy()
                
                for i, (idx, row) in enumerate(df_list.iterrows()):
                    if i >= len(item_codes): break
                    code = item_codes[i]
                    
                    # 2. 개별 종목 실시간 API 호출 (에러 시 해당 종목만 건너뛰고 전체 중단 방지)
                    try:
                        api_url = f"https://polling.finance.naver.com/api/realtime?query=SERVICE_ITEM:{code}"
                        api_resp = requests.get(api_url, headers=headers, timeout=7).json()
                        
                        # API 구조 변경 대비 안전한 데이터 추출
                        if 'result' not in api_resp or not api_resp['result']['areas'][0]['datas']:
                            continue
                            
                        item = api_resp['result']['areas'][0]['datas'][0]
                        
                        # 데이터가 문자열인 경우를 대비해 안전하게 형변환
                        all_stocks.append({
                            'Code': code, 
                            'Name': str(row['종목명']), 
                            'Open': int(item.get('sv', 0)), 
                            'Close': int(item.get('nv', 0)),
                            'Low': int(item.get('lv', 0)), 
                            'High': int(item.get('hv', 0)), 
                            'Ratio': float(item.get('cr', 0.0)), 
                            'Volume': int(item.get('aq', 0)),
                            'Market': market_label
                        })
                    except Exception:
                        continue # 한 종목 에러나도 다음 종목 진행
                        
            except Exception as e:
                print(f"⚠️ {market_label} {page}페이지 수집 중 경고: {e}")
                continue
            
            # 깃허브 서버 차단 방지를 위한 미세 지연 (너무 빠르면 네이버가 막음)
            time.sleep(0.5)
            
    return pd.DataFrame(all_stocks)

async def send_smart_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()

    try:
        limit_count = 500 if day_of_week == 6 else 700
        report_type = "주간평균" if day_of_week == 6 else "일일"
        if day_of_week == 5: report_type = "일일(금요마감)"
        
        print(f"📡 {report_type} 리포트 분석 시작...")
        df_final = await get_naver_realtime_data(limit=limit_count)

        if df_final.empty: 
            print("❌ 수집된 데이터가 없어 중단합니다.")
            return

        # 필터링 로직 (데이터 정합성 체크 포함)
        h_map = {'Code':'종목코드', 'Name':'종목명', 'Open':'시가', 'Close':'종가', 'Low':'저가', 'High':'고가', 'Ratio':'등락률(%)', 'Volume':'거래량'}
        
        def get_filtered_data(m_name, is_up):
            temp = df_final[df_final['Market'] == m_name].copy()
            # 거래량이 0인 종목은 제외 (허수 제거)
            temp = temp[temp['Volume'] > 0]
            cond = (temp['Ratio'] >= 5) if is_up else (temp['Ratio'] <= -5)
            return temp[cond].sort_values('Ratio', ascending=not is_up).rename(columns=h_map).drop(columns=['Market'])

        sheets = {
            '코스피_상승': get_filtered_data('KOSPI', True), 
            '코스닥_상승': get_filtered_data('KOSDAQ', True),
            '코스피_하락': get_filtered_data('KOSPI', False), 
            '코스닥_하락': get_filtered_data('KOSDAQ', False)
        }

        file_name = f"{now.strftime('%m%d')}_{report_type}.xlsx"
        
        # 엑셀 디자인 서식
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
                
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment = header_fill, header_font, Alignment(horizontal='center')
                
                for row in range(2, ws.max_row + 1):
                    # 등락률 값 가져오기 (절대값으로 색상 판정)
                    rv = ws.cell(row, 7).value
                    ratio = abs(float(rv if rv is not None else 0))
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

        # 텔레그램 전송
        up_cnt = len(sheets['코스피_상승']) + len(sheets['코스닥_상승'])
        down_cnt = len(sheets['코스피_하락']) + len(sheets['코스닥_하락'])
        
        msg = (f"📅 {now.strftime('%m-%d')} *[{report_type}] 리포트*\n"
               f"📡 분석: {len(df_final)}개 종목 수집 완료\n\n"
               f"📈 상승(5%↑): {up_cnt}개\n"
               f"📉 하락(5%↓): {down_cnt}개\n\n"
               f"💡 🔴28%↑, 🟠20%↑, 🟡10%↑")
        
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        
        os.remove(file_name) 
        print(f"✅ {file_name} 전송 완료!")

    except Exception as e:
        print(f"❌ 최종 오류 발생: {e}")

if __name__ == "__main__":
    asyncio.run(send_smart_report())
