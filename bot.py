import os, pandas as pd, asyncio, datetime, requests, time
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [1. 설정 정보]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_official_data(target_date=None):
    """정부 API: 특정 날짜 데이터 수집"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '4000', 'resultType': 'json'}
    if target_date:
        params['basDt'] = target_date.replace('-', '')

    try:
        res = requests.get(url, params=params, timeout=30).json()
        items = res['response']['body']['items'].get('item', [])
        if not items:
            params['serviceKey'] = RAW_KEY
            res = requests.get(url, params=params, timeout=30).json()
            items = res['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)
    except:
        return pd.DataFrame()

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    day_of_week = now.weekday() # 0:월, 6:일

    df = pd.DataFrame()
    analysis_info = ""

    # 1. 요일별 데이터 수집 및 모드 결정
    if day_of_week == 0: # [월요일 아침] 지난주 주간 평균 분석
        mode_name = "주간평균리포트"
        weekly_dfs = []
        for i in range(3, 8): # 월요일 기준 3일전(금)~7일전(월)
            target_dt = (now - datetime.timedelta(days=i)).strftime('%Y%m%d')
            df_day = get_official_data(target_dt)
            if not df_day.empty:
                for col in ['mkp', 'clpr', 'lopr', 'hipr', 'fltRt', 'trqu']:
                    df_day[col] = pd.to_numeric(df_day[col], errors='coerce').fillna(0)
                weekly_dfs.append(df_day)
        
        if weekly_dfs:
            full_df = pd.concat(weekly_dfs)
            df = full_df.groupby(['itmsNm', 'mrktCtg', 'srtnCd']).agg({
                'mkp': 'first', 'clpr': 'last', 'lopr': 'min', 'hipr': 'max', 
                'fltRt': 'mean', 'trqu': 'mean'
            }).reset_index()
            df.columns = ['종목명', '시장', '종목코드', '시가', '종가', '저가', '고가', '등락률(%)', '거래량']
            analysis_info = f"지난주 평일 평균 ({(now - datetime.timedelta(days=7)).strftime('%m%d')}~{(now - datetime.timedelta(days=3)).strftime('%m%d')})"
    
    else: # [화~토 아침] 일일 마감 분석 (전일 데이터 자동 탐색)
        mode_name = "일일마감리포트"
        for i in range(1, 8):
            search_date = (now - datetime.timedelta(days=i)).strftime('%Y%m%d')
            raw = get_official_data(search_date)
            if not raw.empty:
                df = pd.DataFrame()
                df['시장'] = raw['mrktCtg']; df['종목코드'] = raw['srtnCd']; df['종목명'] = raw['itmsNm']
                df['시가'] = pd.to_numeric(raw['mkp']).fillna(0); df['종가'] = pd.to_numeric(raw['clpr']).fillna(0)
                df['저가'] = pd.to_numeric(raw['lopr']).fillna(0); df['고가'] = pd.to_numeric(raw['hipr']).fillna(0)
                df['등락률(%)'] = pd.to_numeric(raw['fltRt']).fillna(0)
                df['거래량'] = pd.to_numeric(raw['trqu']).fillna(0)
                analysis_info = f"마감 기준일: {search_date}"
                break
            time.sleep(0.1)

    if df.empty: return

    # 2. 시트 분류 함수
    def filter_data(market, is_up):
        m_cond = df['시장'].str.contains(market)
        r_cond = (df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5)
        return df[m_cond & r_cond].copy().sort_values('등락률(%)', ascending=not is_up)

    sheets_data = {
        '코스피_상승': filter_data('KOSPI', True), '코스피_하락': filter_data('KOSPI', False),
        '코스닥_상승': filter_data('KOSDAQ', True), '코스닥_하락': filter_data('KOSDAQ', False)
    }

    # 3. 엑셀 생성 및 디자인 (열 너비 및 강조 적용)
    file_name = f"{now.strftime('%m%d')}_{mode_name}.xlsx"
    header_fill = PatternFill("solid", fgColor="444444")
    font_white = Font(color="FFFFFF", bold=True)
    colors = {'red': PatternFill("solid", "FF0000"), 'orange': PatternFill("solid", "FFCC00"), 'yellow': PatternFill("solid", "FFFF00")}
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in sheets_data.items():
            data.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # [수정] 열 너비 설정
            ws.column_dimensions['A'].width = 25 # 종목명
            ws.column_dimensions['B'].width = 12 # 시장
            ws.column_dimensions['C'].width = 12 # 종목코드
            for col in ['D','E','F','G']: ws.column_dimensions[ws.cell(1,ws.get_column_letter(data.columns.get_loc(col)+1).column).column_letter].width = 15
            ws.column_dimensions['H'].width = 12 # 등락률
            ws.column_dimensions['I'].width = 25 # 거래량 (중요: 너비 확대)

            for r in range(1, ws.max_row + 1):
                for c in range(1, 10):
                    cell = ws.cell(r, c)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    if r == 1:
                        cell.fill, cell.font = header_fill, font_white
                    else:
                        if c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                        if c == 8: # 등락률 소수점 및 색상 강조
                            cell.number_format = '0.00'
                            val = abs(float(cell.value or 0))
                            target_cell = ws.cell(r, 1) # 종목명 강조
                            if val >= 25: target_cell.fill, target_cell.font = colors['red'], font_white
                            elif val >= 20: target_cell.fill = colors['orange']
                            elif val >= 10: target_cell.fill = colors['yellow']

    # 4. 전송 및 메시지 디자인
    total_up = len(sheets_data['코스피_상승']) + len(sheets_data['코스닥_상승'])
    total_down = len(sheets_data['코스피_하락']) + len(sheets_data['코스닥_하락'])

    async with bot:
        msg = (f"📊 모드: {mode_name}\n"
               f"🔍 {analysis_info}\n\n"
               f"📈 상승(5%↑): {total_up}개\n"
               f"📉 하락(5%↓): {total_down}개\n"
               f"💡 🟡10%↑ 🟠20%↑ 🔴25%↑")
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
