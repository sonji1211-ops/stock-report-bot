import os, pandas as pd, asyncio, datetime, requests
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_realtime_naver():
    """[장중용] 네이버 실시간 API 데이터 수집 (시/고/저가는 0으로 처리될 수 있음)"""
    results = []
    for sosok in [0, 1]:
        market = "KOSPI" if sosok == 0 else "KOSDAQ"
        for page in range(1, 45):
            url = f"https://m.stock.naver.com/api/stocks/marketValue/{sosok}?page={page}&pageSize=50"
            try:
                res = requests.get(url, timeout=10).json().get('result', [])
                if not res: break
                for item in res:
                    results.append({
                        '시장': market, '종목코드': item['itemCode'], '종목명': item['stockName'],
                        '시가': 0, '종가': int(item['closePrice'].replace(',', '')), 
                        '저가': 0, '고가': 0, '등락률(%)': float(item['fluctuationsRatio']),
                        '거래량': int(item['accumulatedTradingVolume'].replace(',', ''))
                    })
            except: break
    return pd.DataFrame(results)

def get_official_gov():
    """[마감/주간용] 정부 공식 API 데이터 수집"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '3000', 'resultType': 'json'}
    try:
        res = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        df_raw = pd.DataFrame(res)
    except:
        params['serviceKey'] = RAW_KEY
        res = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        df_raw = pd.DataFrame(res)
    
    df = pd.DataFrame()
    df['시장'] = df_raw['mrktCtg']
    df['종목코드'] = df_raw['srtnCd']
    df['종목명'] = df_raw['itmsNm']
    df['시가'] = pd.to_numeric(df_raw['mkp'])
    df['종가'] = pd.to_numeric(df_raw['clpr'])
    df['저가'] = pd.to_numeric(df_raw['lopr'])
    df['고가'] = pd.to_numeric(df_raw['hipr'])
    df['등락률(%)'] = pd.to_numeric(df_raw['fltRt'])
    df['거래량'] = pd.to_numeric(df_raw['trqu'])
    return df

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.now()
    day_of_week = now.weekday() # 6은 일요일
    hour = now.hour

    # 1. 모드 결정 및 데이터 수집
    if day_of_week == 6:
        mode_name = "주간평균(공식)"
        df = get_official_gov() # 일요일은 주간 분석용 공식 데이터
        analysis_info = "한 주간의 변동성 분석"
    elif 9 <= hour < 16:
        mode_name = "장중실시간(네이버)"
        df = get_realtime_naver()
        analysis_info = "현재가 기준 실시간 전수조사"
    else:
        mode_name = "일일마감(공식)"
        df = get_official_gov()
        analysis_info = "정부 데이터 공식 종가 기준"

    if df.empty: return

    # 2. 시트 분류
    def get_sheet(market, is_up):
        cond = (df['시장'] == market) & ((df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5))
        return df[cond].sort_values('등락률(%)', ascending=not is_up).drop(columns=['시장'])

    sheets_data = {
        '코스피_상승': get_sheet('KOSPI', True), '코스피_하락': get_sheet('KOSPI', False),
        '코스닥_상승': get_sheet('KOSDAQ', True), '코스닥_하락': get_sheet('KOSDAQ', False)
    }

    # 3. 엑셀 생성 및 디자인 반영
    file_name = f"{now.strftime('%m%d')}_{mode_name}.xlsx"
    h_fill = PatternFill("solid", fgColor="444444")
    f_white = Font(color="FFFFFF", bold=True)
    colors = {'red': PatternFill("solid", "FF0000"), 'orange': PatternFill("solid", "FFCC00"), 'yellow': PatternFill("solid", "FFFF00")}
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in sheets_data.items():
            data.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # 너비 및 헤더 스타일
            ws.column_dimensions['B'].width = 12 # 종목코드
            ws.column_dimensions['C'].width = 28 # 종목명
            for c_idx in [4,5,6,7,9]: ws.column_dimensions[chr(64+c_idx)].width = 15
            
            for r in range(1, ws.max_row + 1):
                for c in range(1, 10):
                    cell = ws.cell(r, c)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    if r == 1:
                        cell.fill, cell.font = h_fill, f_white
                    else:
                        if c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                        if c == 8: # 등락률
                            cell.number_format = '0.00'
                            val = abs(float(cell.value or 0))
                            target = ws.cell(r, 3) # 종목명 색상
                            if val >= 25: target.fill, target.font = colors['red'], f_white
                            elif val >= 20: target.fill = colors['orange']
                            elif val >= 10: target.fill = colors['yellow']

    # 4. 텔레그램 발송
    total_up = len(sheets_data['코스피_상승']) + len(sheets_data['코스닥_상승'])
    total_down = len(sheets_data['코스피_하락']) + len(sheets_data['코스닥_하락'])

    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d %H:%M')}\n"
               f"📊 모드: {mode_name}\n"
               f"🔍 분석: {analysis_info}\n\n"
               f"🔺 상승(5%↑): {total_up}개\n"
               f"🔻 하락(5%↓): {total_down}개\n\n"
               f"💡 🟡10%↑ 🟠20%↑ 🔴25%↑")
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
