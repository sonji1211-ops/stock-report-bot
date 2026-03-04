import os, pandas as pd, asyncio, datetime, requests, time
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_realtime_naver():
    results = []
    headers = {"User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1", "Referer": "https://m.stock.naver.com/"}
    for sosok in [0, 1]:
        market = "KOSPI" if sosok == 0 else "KOSDAQ"
        for page in range(1, 45):
            url = f"https://m.stock.naver.com/api/stocks/marketValue/{sosok}?page={page}&pageSize=50"
            try:
                res = requests.get(url, headers=headers, timeout=10).json().get('result', [])
                if not res: break
                for item in res:
                    results.append({
                        '시장': market, '종목코드': item['itemCode'], '종목명': item['stockName'],
                        '시가': int(str(item.get('openPrice', '0')).replace(',', '') or 0),
                        '종가': int(str(item.get('closePrice', '0')).replace(',', '') or 0),
                        '저가': int(str(item.get('lowPrice', '0')).replace(',', '') or 0),
                        '고가': int(str(item.get('highPrice', '0')).replace(',', '') or 0),
                        '등락률(%)': float(item.get('fluctuationsRatio', 0)),
                        '거래량': int(str(item.get('accumulatedTradingVolume', '0')).replace(',', '') or 0)
                    })
            except: break
    return pd.DataFrame(results)

def get_official_gov():
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '4000', 'resultType': 'json'}
    try:
        r = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        if not r: 
            params['serviceKey'] = RAW_KEY
            r = requests.get(url, params=params, timeout=30).json()['response']['body']['items'].get('item', [])
        df_raw = pd.DataFrame(r)
        df = pd.DataFrame()
        df['시장'] = df_raw['mrktCtg']; df['종목코드'] = df_raw['srtnCd']; df['종목명'] = df_raw['itmsNm']
        df['시가'] = pd.to_numeric(df_raw['mkp']).fillna(0).astype(int)
        df['종가'] = pd.to_numeric(df_raw['clpr']).fillna(0).astype(int)
        df['저가'] = pd.to_numeric(df_raw['lopr']).fillna(0).astype(int)
        df['고가'] = pd.to_numeric(df_raw['hipr']).fillna(0).astype(int)
        df['등락률(%)'] = pd.to_numeric(df_raw['fltRt']).fillna(0).astype(float)
        df['거래량'] = pd.to_numeric(df_raw['trqu']).fillna(0).astype(int)
        return df
    except: return pd.DataFrame()

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    day_of_week, hour = now.weekday(), now.hour

    # 1. 모드 판정 및 수집
    if day_of_week == 6: mode, df = "주간분석", get_official_gov()
    elif day_of_week < 5 and (9 <= hour < 16): mode, df = "장중실시간", get_realtime_naver()
    else: mode, df = "일일마감", get_official_gov()

    if df.empty: return

    # 2. 엑셀 파일 생성
    file_name = f"{now.strftime('%m%d_%H%M')}_{mode}.xlsx"
    h_fill = PatternFill("solid", fgColor="444444")
    f_white_b = Font(color="FFFFFF", bold=True)
    colors = {'red': PatternFill("solid", "FF0000"), 'orange': PatternFill("solid", "FFCC00"), 'yellow': PatternFill("solid", "FFFF00")}
    border = Side(style='thin')

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m_name in ['KOSPI', 'KOSDAQ']:
            for is_up in [True, False]:
                cond = (df['시장'] == m_name) & ((df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5))
                sub_df = df[cond].sort_values('등락률(%)', ascending=not is_up).drop(columns=['시장'])
                sheet_label = f"{m_name}_{'상승' if is_up else '하락'}"
                sub_df.to_excel(writer, sheet_name=sheet_label, index=False)
                
                ws = writer.sheets[sheet_label]
                # 컬럼 너비 설정 (거래량 H열 25로 대폭 확대)
                for i, w in enumerate([12, 30, 14, 14, 14, 14, 12, 25], 1):
                    ws.column_dimensions[ws.cell(1, i).column_letter].width = w

                for r in range(1, ws.max_row + 1):
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = Border(left=border, right=border, top=border, bottom=border)
                        if r == 1:
                            cell.fill, cell.font = h_fill, f_white_b
                        else:
                            if c in [3,4,5,6,8]: cell.number_format = '#,##0'
                            if c == 7: # 등락률에 따른 종목명 강조
                                val = abs(float(cell.value or 0))
                                target = ws.cell(r, 2)
                                if val >= 25: target.fill, target.font = colors['red'], f_white_b
                                elif val >= 20: target.fill = colors['orange']
                                elif val >= 10: target.fill = colors['yellow']

    # 3. 텔레그램 메시지 디자인 및 전송
    up_cnt = len(df[df['등락률(%)'] >= 5])
    down_cnt = len(df[df['등락률(%)'] <= -5])
    
    msg = (f"📂 {now.strftime('%Y-%m-%d %H:%M')} {mode}\n"
           f"━━━━━━━━━━━━━━━━━━\n"
           f"📈 상승(5%↑): {up_cnt}개 종목\n"
           f"📉 하락(5%↓): {down_cnt}개 종목\n"
           f"━━━━━━━━━━━━━━━━━━\n"
           f"💡 가독성: 거래량 너비 확대 완료\n"
           f"💡 강조: 🟡10%↑ 🟠20%↑ 🔴25%↑")

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
