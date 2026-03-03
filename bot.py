import os, pandas as pd, asyncio, datetime, requests, time
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# 공공데이터포털 인증키 (인코딩된 키를 안전하게 디코딩하여 사용)
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_krx_all_stocks():
    """공공데이터포털 API: 전 종목 시세를 단 1회 호출로 수집"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockPriceInfoService/getStockPriceInfo'
    
    # 3,000개 종목을 한 번에 요청 (전수조사)
    params = {
        'serviceKey': SERVICE_KEY,
        'numOfRows': '3000',
        'pageNo': '1',
        'resultType': 'json'
    }

    try:
        response = requests.get(url, params=params, timeout=30)
        if response.status_code != 200:
            print(f"🚨 서버 응답 에러: {response.status_code}")
            return None
            
        res_data = response.json()
        
        # 데이터 구조 진입 및 유효성 검사
        if 'response' in res_data and 'body' in res_data['response']:
            items_data = res_data['response']['body']['items'].get('item', [])
            if not items_data:
                print("⚠️ 수집된 종목이 없습니다. API 승인 대기 중인지 확인하세요.")
                return None
            return pd.DataFrame(items_data)
        else:
            print(f"🚨 API 응답 오류: {res_data.get('header', {}).get('resultMsg')}")
            return None
            
    except Exception as e:
        print(f"🚨 API 호출 중 예외 발생: {e}")
        return None

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    print("📡 [정부 데이터] 전 종목 전수조사 시작...")
    df_raw = get_krx_all_stocks()
    
    if df_raw is None or df_raw.empty:
        print("❌ 분석을 중단합니다. (데이터 없음)")
        return

    # 1. 데이터 표준화 (정부 API -> 리포트용)
    # srtnCd: 단축코드, itmsNm: 종목명, mrktCtg: 시장구분, fltRt: 등락률
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

    # 3. 엑셀 생성
    file_name = f"{now.strftime('%m%d')}_국내증시_전수조사.xlsx"
    h_fill = PatternFill("solid", "444444")
    f_white = Font(color="FFFFFF", bold=True)
    p_red, p_ora, p_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', trend_name := '하락']:
                cond = (final_df['Ratio'] > 0) if trend == '상승' else (final_df['Ratio'] < 0)
                sub = final_df[(final_df['Market'] == m) & cond].drop(columns=['Market'])
                
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 디자인 적용
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = h_fill, f_white, Alignment(horizontal='center'), border
                
                for r in range(2, ws.max_row + 1):
                    ratio_val = abs(float(ws.cell(r, 7).value or 0))
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment, cell.border = Alignment(horizontal='center'), border
                        if c in [3,4,5,6,8]: cell.number_format = '#,##0'
                        if c == 7: cell.number_format = '0.00'
                        if c == 2: # 강조색
                            if ratio_val >= 28: cell.fill, cell.font = p_red, f_white
                            elif ratio_val >= 20: cell.fill = p_ora
                            elif ratio_val >= 10: cell.fill = p_yel
                ws.column_dimensions['B'].width = 18

    # 4. 텔레그램 발송
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 국내증시 리포트\n\n"
               f"📊 조사대상: {len(df)}개 전 종목\n"
               f"⚡ 변동폭(5%↑↓): {len(final_df)}개\n\n"
               f"✅ 유령 번호 0%, 정부 공식 데이터 사용")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    
    if os.path.exists(file_name):
        os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
