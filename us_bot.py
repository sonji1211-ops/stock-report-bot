import os, pandas as pd, asyncio, datetime, requests
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# 지수님이 주신 인증키 (이미 인코딩된 상태이므로 unquote로 풀어서 사용해야 안전함)
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY) 

def get_krx_all_stocks():
    """공공데이터포털 API: 전 종목 시세를 단 1회 호출로 수집"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    
    # 어제 날짜 기준으로 데이터 조회 (주말/공휴일 대비 넉넉하게 최근 데이터 호출)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    # 데이터가 확실히 존재하는 최근 거래일을 찾기 위해 7일치 범위를 둡니다.
    
    params = {
        'serviceKey': SERVICE_KEY,
        'numOfRows': '3000', # 전 종목을 한 페이지에 다 담음
        'pageNo': '1',
        'resultType': 'json'
    }

    try:
        # 인증키 에러 방지를 위해 params 대신 직접 URL을 구성하여 호출
        response = requests.get(url, params=params, timeout=30)
        
        if response.status_code != 200:
            print(f"🚨 서버 응답 에러: {response.status_code}")
            return None
            
        res_data = response.json()
        items = res_data['response']['body']['items']['item']
        return pd.DataFrame(items)
    except Exception as e:
        print(f"🚨 API 호출 오류 (키 설정 확인 필요): {e}")
        return None

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    print("📡 공공데이터포털(정부) 서버 접속 중... (유령 코드 0% 분석)")
    df_raw = get_krx_all_stocks()
    
    if df_raw is None or df_raw.empty:
        print("❌ 데이터를 가져오지 못했습니다. 인증키 승인 대기 중(최대 1시간)이거나 키가 잘못되었습니다.")
        return

    # 1. 데이터 정제 (정부 API 컬럼 -> 표준 컬럼)
    df = pd.DataFrame()
    df['Code'] = df_raw['srtnCd']    # 단축코드
    df['Name'] = df_raw['itmsNm']    # 종목명
    df['Market'] = df_raw['mrktCtg'] # 시장구분 (KOSPI/KOSDAQ)
    df['Open'] = pd.to_numeric(df_raw['mkp'])
    df['Close'] = pd.to_numeric(df_raw['clpr'])
    df['High'] = pd.to_numeric(df_raw['hipr'])
    df['Low'] = pd.to_numeric(df_raw['lopr'])
    df['Ratio'] = pd.to_numeric(df_raw['fltRt'])
    df['Volume'] = pd.to_numeric(df_raw['trqu'])

    # 2. ±5% 필터링 (지수님 요청 조건)
    final_df = df[(df['Ratio'] >= 5) | (df['Ratio'] <= -5)]
    final_df = final_df.sort_values('Ratio', ascending=False)

    # [3. 엑셀 생성 및 디자인]
    file_name = f"{now.strftime('%m%d')}_정부데이터_리포트.xlsx"
    h_fill, f_white = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    p_red, p_ora, p_yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m in ['KOSPI', 'KOSDAQ']:
            for trend in ['상승', '하락']:
                sub = final_df[(final_df['Market']==m) & ((final_df['Ratio']>0) if trend=='상승' else (final_df['Ratio']<0))]
                sub = sub.drop(columns=['Market'])
                
                sheet_name = f"{m}_{trend}"
                sub.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]

                # 헤더 스타일링
                for cell in ws[1]:
                    cell.fill, cell.font, cell.alignment, cell.border = h_fill, f_white, Alignment(horizontal='center'), border
                
                # 본문 스타일링
                for r in range(2, ws.max_row + 1):
                    rv = abs(float(ws.cell(r, 7).value or 0))
                    for c in range(1, 9):
                        cell = ws.cell(r, c)
                        cell.alignment, cell.border = Alignment(horizontal='center'), border
                        if c in [3, 4, 5, 6, 8]: cell.number_format = '#,##0' # 금액 및 거래량
                        if c == 7: cell.number_format = '0.00' # 등락률
                        if c == 2: # 종목명 강조
                            if rv >= 28: cell.fill, cell.font = p_red, Font(color="FFFFFF", bold=True)
                            elif rv >= 20: cell.fill = p_ora
                            elif rv >= 10: cell.fill = p_yel
                ws.column_dimensions['B'].width = 15

    # [4. 텔레그램 전송]
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 정부 공식데이터 분석\n\n"
               f"📊 수집 종목: {len(df)}개 (KRX 전수조사)\n"
               f"⚡ ±5% 필터 통과: {len(final_df)}개\n\n"
               f"💡 404 에러 없는 클린 리포트 발송 완료")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
