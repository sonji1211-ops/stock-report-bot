import os, pandas as pd, asyncio, datetime, requests
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_official_krx():
    """정부 API를 이용해 전 종목 최신 시세를 수집 (거래량 포함)"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {
        'serviceKey': SERVICE_KEY,
        'numOfRows': '3000',
        'pageNo': '1',
        'resultType': 'json'
    }
    try:
        response = requests.get(url, params=params, timeout=30)
        items = response.json()['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)
    except:
        params['serviceKey'] = RAW_KEY
        response = requests.get(url, params=params, timeout=30)
        items = response.json()['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.now()
    
    print("📡 거래량 포함 데이터 수집 중...")
    df_raw = get_official_krx()
    
    if df_raw.empty: return

    # 1. 데이터 정리 및 한글 항목명 변경 (거래량 추가)
    df = pd.DataFrame()
    df['시장'] = df_raw['mrktCtg']
    df['종목명'] = df_raw['itmsNm']
    df['현재가'] = pd.to_numeric(df_raw['clpr'])
    df['등락(수치)'] = pd.to_numeric(df_raw['vs'])
    df['등락률(%)'] = pd.to_numeric(df_raw['fltRt'])
    df['거래량'] = pd.to_numeric(df_raw['trqu']) # [추가] 거래량 데이터 변환

    # 2. 통계 계산 (전 종목 기준)
    up_count = len(df[df['등락률(%)'] > 0])
    down_count = len(df[df['등락률(%)'] < 0])
    even_count = len(df[df['등락률(%)'] == 0])
    
    # 3. ±5% 필터링 종목 준비
    plus_5 = df[df['등락률(%)'] >= 5].sort_values('등락률(%)', ascending=False)
    # 하락 정렬: 가장 많이 떨어진 순서 (-30%부터 보임)
    minus_5 = df[df['등락률(%)'] <= -5].sort_values('등락률(%)', ascending=True)

    file_name = f"{now.strftime('%m%d')}_국내증시_전수조사.xlsx"
    
    # 4. 엑셀 생성 및 디자인
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        plus_5.to_excel(writer, sheet_name='급상승(5%↑)', index=False)
        minus_5.to_excel(writer, sheet_name='급하락(5%↓)', index=False)
        
        for sheet in ['급상승(5%↑)', '급하락(5%↓)']:
            ws = writer.sheets[sheet]
            header_fill = PatternFill("solid", "444444")
            white_font = Font(color="FFFFFF", bold=True)
            border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

            # 열 너비 설정 (종목명 짤림 방지 및 거래량 칸 확보)
            ws.column_dimensions['A'].width = 10 # 시장
            ws.column_dimensions['B'].width = 28 # 종목명 (충분히 확장)
            ws.column_dimensions['C'].width = 15 # 현재가
            ws.column_dimensions['D'].width = 12 # 등락(수치)
            ws.column_dimensions['E'].width = 12 # 등락률(%)
            ws.column_dimensions['F'].width = 18 # 거래량

            for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    if cell.row == 1:
                        cell.fill, cell.font = header_fill, white_font
                    
                    # [포맷팅] 현재가, 수치, 거래량에 천 단위 콤마 추가
                    if cell.column in [3, 4, 6]:
                        cell.number_format = '#,##0'
                    # 등락률 소수점 표시
                    if cell.column == 5:
                        cell.number_format = '0.00'

    # 5. 텔레그램 메시지 발송
    msg = (
        f"📅 {now.strftime('%Y-%m-%d')} 국장 전수조사\n"
        f"━━━━━━━━━━━━━━━\n"
        f"📊 전종목 등락 요약\n"
        f"  • 상승 🔺: {up_count}개\n"
        f"  • 하락 🔻: {down_count}개\n"
        f"  • 보합 ➖: {even_count}개\n"
        f"━━━━━━━━━━━━━━━\n"
        f"🔥 급변동 (±5% 이상)\n"
        f"  • 급상승 🚀: {len(plus_5)}개\n"
        f"  • 급하락 📉: {len(minus_5)}개\n"
        f"━━━━━━━━━━━━━━━\n"
        f"✅ 거래량 항목 추가 및 정렬 완료"
    )

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
