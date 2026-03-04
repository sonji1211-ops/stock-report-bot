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
    """[장중용] 네이버 실시간 API - 시/고/저가까지 정밀 수집"""
    results = []
    for sosok in [0, 1]:
        market = "KOSPI" if sosok == 0 else "KOSDAQ"
        for page in range(1, 45):
            url = f"https://m.stock.naver.com/api/stocks/marketValue/{sosok}?page={page}&pageSize=50"
            try:
                res = requests.get(url, timeout=10).json().get('result', [])
                if not res: break
                for item in res:
                    # 실시간 모드에서 누락 없는 데이터 맵핑
                    results.append({
                        '시장': market,
                        '종목코드': item['itemCode'],
                        '종목명': item['stockName'],
                        '시가': int(item.get('openPrice', '0').replace(',', '')),
                        '종가': int(item['closePrice'].replace(',', '')),
                        '저가': int(item.get('lowPrice', '0').replace(',', '')),
                        '고가': int(item.get('highPrice', '0').replace(',', '')),
                        '등락률(%)': float(item['fluctuationsRatio']),
                        '거래량': int(item['accumulatedTradingVolume'].replace(',', ''))
                    })
            except: break
    return pd.DataFrame(results)

def get_official_gov():
    """[마감/주간용] 정부 공식 API"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '4000', 'resultType': 'json'}
    try:
        res = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        df_raw = pd.DataFrame(res)
    except:
        params['serviceKey'] = RAW_KEY
        res = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        df_raw = pd.DataFrame(res)
    
    if df_raw.empty: return pd.DataFrame()

    df = pd.DataFrame()
    df['시장'] = df_raw['mrktCtg']
    df['종목코드'] = df_raw['srtnCd']
    df['종목명'] = df_raw['itmsNm']
    df['시가'] = pd.to_numeric(df_raw['mkp'], errors='coerce').fillna(0)
    df['종가'] = pd.to_numeric(df_raw['clpr'], errors='coerce').fillna(0)
    df['저가'] = pd.to_numeric(df_raw['lopr'], errors='coerce').fillna(0)
    df['고가'] = pd.to_numeric(df_raw['hipr'], errors='coerce').fillna(0)
    df['등락률(%)'] = pd.to_numeric(df_raw['fltRt'], errors='coerce').fillna(0)
    df['거래량'] = pd.to_numeric(df_raw['trqu'], errors='coerce').fillna(0)
    return df

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.now()
    day_of_week = now.weekday()
    hour = now.hour

    # 1. 모드 결정 (장중 09:00 ~ 15:40 실시간 모드)
    if day_of_week < 5 and (9 <= hour < 16):
        mode_name = "장중실시간"
        df = get_realtime_naver()
        analysis_info = "🚀 네이버 실시간 체결 데이터"
    else:
        mode_name = "공식마감"
        df = get_official_gov()
        analysis_info = "🏛️ 정부 공식 확정 데이터"

    if df.empty: 
        print("데이터를 가져오지 못했습니다.")
        return

    # 2. 시트 분류 및 정렬 (하락은 -30%가 맨 위로)
    def get_sheet(market, is_up):
        cond = (df['시장'] == market) & ((df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5))
        target_df = df[cond].copy()
        sort_order = not is_up # 상승은 내림차순, 하락은 오름차순
        return target_df.sort_values('등락률(%)', ascending=sort_order).drop(columns=['시장'])

    sheets_data = {
        '코스피_상승': get_sheet('KOSPI', True), '코스피_하락': get_sheet('KOSPI', False),
        '코스닥_상승': get_sheet('KOSDAQ', True), '코스닥_하락': get_sheet('KOSDAQ', False)
    }

    # 3. 엑셀 생성 및 디자인 (열 순서 고정)
    file_name = f"{now.strftime('%m%d_%H%M')}_{mode_name}.xlsx"
    h_fill = PatternFill("solid", fgColor="444444")
    f_white = Font(color="FFFFFF", bold=True)
    colors = {'red': PatternFill("solid", "FF0000"), 'orange': PatternFill("solid", "FFCC00"), 'yellow': PatternFill("solid", "FFFF00")}
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in sheets_data.items():
            data.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # 너비 설정
            ws.column_dimensions['A'].width = 12 # 종목코드
            ws.column_dimensions['B'].width = 25 # 종목명
            for col in ['C','D','E','F','H']: ws.column_dimensions[ws.cell(1, data.columns.get_loc({'시가':'시가','종가':'종가','저가':'저가','고가':'고가','거래량':'거래량'}.get(ws.cell(1,3).value, '거래량'))+1).column_letter].width = 14
            ws.column_dimensions['G'].width = 12 # 등락률

            for r in range(1, ws.max_row + 1):
                for c in range(1, 9): # A~H열
                    cell = ws.cell(r, c)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    if r == 1:
                        cell.fill, cell.font = h_fill, f_white
                    else:
                        # 숫자 포맷 (시/종/저/고/거래량)
                        if c in [3, 4, 5, 6, 8]: cell.number_format = '#,##0'
                        if c == 7: # 등락률 소수점 및 색상
                            cell.number_format = '0.00'
                            val = abs(float(cell.value or 0))
                            nm_cell = ws.cell(r, 2) # 종목명 셀
                            if val >= 25: nm_cell.fill, nm_cell.font = colors['red'], f_white
                            elif val >= 20: nm_cell.fill = colors['orange']
                            elif val >= 10: nm_cell.fill = colors['yellow']

    # 4. 텔레그램 발송
    async with bot:
        caption = (f"📅 {now.strftime('%Y-%m-%d %H:%M')}\n"
                   f"📊 모드: {mode_name}\n"
                   f"🔍 {analysis_info}\n\n"
                   f"💡 🟡10%↑ 🟠20%↑ 🔴25%↑")
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=caption)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
