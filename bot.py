import os, pandas as pd, asyncio, datetime, requests
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [설정]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_official_data(target_date=None):
    """정부 API: 특정 날짜 혹은 최신 데이터 수집"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {
        'serviceKey': SERVICE_KEY,
        'numOfRows': '3000',
        'pageNo': '1',
        'resultType': 'json'
    }
    if target_date:
        params['basDt'] = target_date.replace('-', '') # YYYYMMDD 형식

    try:
        res = requests.get(url, params=params, timeout=30).json()
        items = res['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)
    except:
        params['serviceKey'] = RAW_KEY
        res = requests.get(url, params=params, timeout=30).json()
        items = res['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.now() # 한국 시간 기준 실행 가정
    day_of_week = now.weekday() 

    # 1. 요일별 데이터 수집
    if day_of_week == 6: # [일요일] 주간 분석
        report_type = "주간평균"
        # 주간은 데이터 특성상 최신 영업일 기준 등락 확인 (API 제약상 금요일 데이터 사용)
        df_raw = get_official_data()
        analysis_info = "주간 변동성(최신영업일 기준)"
    else: # [화~토] 일일 분석
        report_type = "일일"
        df_raw = get_official_data()
        analysis_info = "전일자 전수조사"

    if df_raw.empty:
        print("❌ 데이터를 가져오지 못했습니다.")
        return

    # 2. 데이터 정제 및 한글화
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

    # 3. 시트 분류 (코스피/코스닥 x 상승/하락)
    def filter_data(market, is_up):
        m_cond = df['시장'].str.contains(market)
        r_cond = (df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5)
        res = df[m_cond & r_cond].copy()
        return res.sort_values('등락률(%)', ascending=not is_up)

    sheets_data = {
        '코스피_상승': filter_data('KOSPI', True),
        '코스피_하락': filter_data('KOSPI', False),
        '코스닥_상승': filter_data('KOSDAQ', True),
        '코스닥_하락': filter_data('KOSDAQ', False)
    }

    # 4. 엑셀 생성 및 디자인
    file_name = f"{now.strftime('%m%d')}_{report_type}_리포트.xlsx"
    fill_red = PatternFill("solid", fgColor="FF0000")    # 25%↑
    fill_orange = PatternFill("solid", fgColor="FFCC00") # 20%↑
    fill_yellow = PatternFill("solid", fgColor="FFFF00") # 10%↑
    header_fill = PatternFill("solid", fgColor="444444")
    font_white = Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, data in sheets_data.items():
            data.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # 너비 조절
            ws.column_dimensions['B'].width = 12 # 종목코드
            ws.column_dimensions['C'].width = 25 # 종목명
            for col in ['D','E','F','G','I']: ws.column_dimensions[ws.cell(1, data.columns.get_loc({'시가':'시가','종가':'종가','저가':'저가','고가':'고가','거래량':'거래량'}[ws.cell(1, data.columns.get_loc(ws.cell(1,6).value)+1).value if False else '거래량' ])+1).column_letter].width = 15 # 자동너비 대신 수동지정
            ws.column_dimensions['H'].width = 12 # 등락률

            for r in range(1, ws.max_row + 1):
                for c in range(1, 10):
                    cell = ws.cell(r, c)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = border
                    if r == 1:
                        cell.fill, cell.font = header_fill, font_white
                    else:
                        if c in [4, 5, 6, 7, 9]: cell.number_format = '#,##0'
                        if c == 8:
                            cell.number_format = '0.00'
                            val = abs(float(cell.value or 0))
                            target_cell = ws.cell(r, 3) # 종목명 색상
                            if val >= 25: target_cell.fill, target_cell.font = fill_red, font_white
                            elif val >= 20: target_cell.fill = fill_orange
                            elif val >= 10: target_cell.fill = fill_yellow

    # 5. 전송
    total_up = len(sheets_data['코스피_상승']) + len(sheets_data['코스닥_상승'])
    total_down = len(sheets_data['코스피_하락']) + len(sheets_data['코스닥_하락'])

    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} {report_type} 리포트\n"
               f"📊 {analysis_info}\n"
               f"📈 상승(5%↑): {total_up}개 / 📉 하락(5%↓): {total_down}개\n"
               f"💡 🟡10%↑ 🟠20%↑ 🔴25%↑ (종목명 강조)")
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
