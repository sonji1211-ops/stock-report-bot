import os, pandas as pd, asyncio, datetime, requests, time
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# 공공데이터포털 인증키 (지수님 키)
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
# API 서버에 따라 unquote가 필요한 경우가 있어 두 가지를 대비합니다.
SERVICE_KEY = unquote(RAW_KEY)

def get_krx_all_stocks():
    """공공데이터포털 API: 전 종목 시세를 단 1회 호출로 수집"""
    # [최신화된 주소] 404 에러 방지용 공식 주소
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    
    params = {
        'serviceKey': SERVICE_KEY,
        'numOfRows': '3000', # 전 종목 수용
        'pageNo': '1',
        'resultType': 'json'
    }

    try:
        response = requests.get(url, params=params, timeout=30)
        
        if response.status_code != 200:
            print(f"🚨 서버 응답 에러: {response.status_code}")
            # 404나 500 에러 발생 시 키를 unquote 없이 시도 (예외 처리)
            params['serviceKey'] = RAW_KEY
            response = requests.get(url, params=params, timeout=30)
            if response.status_code != 200: return None
            
        res_data = response.json()
        
        if 'response' in res_data and 'body' in res_data['response']:
            items_data = res_data['response']['body']['items'].get('item', [])
            return pd.DataFrame(items_data)
        else:
            print(f"🚨 데이터 구조 오류: {res_data.get('header', {}).get('resultMsg')}")
            return None
            
    except Exception as e:
        print(f"🚨 API 호출 예외: {e}")
        return None

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    print("📡 [정부 공식 데이터] 전 종목 전수조사 시작...")
    df_raw = get_krx_all_stocks()
    
    if df_raw is None or df_raw.empty:
        print("❌ 데이터를 가져오지 못했습니다. (API 키 승인 대기 또는 주소 확인)")
        return

    # 1. 데이터 정제
    df = pd.DataFrame()
    df['Code'] = df_raw['srtnCd']
    df['Name'] = df_raw['itmsNm']
    df['Market'] = df_raw['mrktCtg']
    df['Open'] = pd.to_numeric(df_raw['mkp'], errors='coerce').fillna(0)
    df['Close'] = pd.to_numeric(df_raw['clpr'], errors='coerce').fillna(0)
    df['High'] = pd.to_numeric(df_raw['hipr'], errors='coerce').fillna(0)
    df['Low'] = pd.to_numeric(df_raw['lopr'], errors='coerce').fillna(0)
    df['Ratio'] = pd.to_numeric(df_raw['fltRt'], errors='coerce').fillna(0)
    df['Volume'] = pd.to_numeric(df_raw['trqu'], errors='coerce').fillna(0)

    # 2. ±5% 필터링
    final_df = df[(df['Ratio'] >= 5) | (df['Ratio'] <= -5)]
    final_df = final_df.sort_values('Ratio', ascending=False)

    # 3. 엑셀 파일 생성
    file_name = f"{now.strftime('%m%d')}_국내증시_리포트.xlsx"
    h_fill, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    p_red, p_ora, p_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                cond = (final_df['Ratio'] > 0) if trend == '상승' else (final_df['Ratio'] < 0)
                sub = final_df[(final_df['Market'] == m) & cond].drop(columns=['Market'])
                
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 디자인 입히기
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = h_fill, f_white, Alignment(horizontal='center'), border
                
                for r in range(2, ws.max_row + 1):
                    ratio_val = abs(float(ws.cell(r, 7).value or 0))
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment, cell.border = Alignment(horizontal='center'), border
                        if c in [3,4,5,6,8]: cell.number_format = '#,##0'
                        if c == 7: cell.number_format = '0.00'
                        if c == 2:
                            if ratio_val >= 28: cell.fill, cell.font = p_red, f_white
                            elif ratio_val >= 20: cell.fill = p_ora
                            elif ratio_val >= 10: cell.fill = p_yel
                ws.column_dimensions['B'].width = 18

    # 4. 텔레그램 발송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 국내증시 리포트\n\n"
               f"📊 조사대상: {len(df)}개 전 종목\n"
               f"⚡ 변동폭(5%↑↓): {len(final_df)}개\n\n"
               f"✅ 유령 번호 없음 / 정부 데이터 기반")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
