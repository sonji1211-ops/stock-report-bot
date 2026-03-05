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
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '4000', 'resultType': 'json'}
    if target_date: params['basDt'] = target_date.replace('-', '')
    try:
        res = requests.get(url, params=params, timeout=30).json()
        items = res['response']['body']['items'].get('item', [])
        if not items:
            params['serviceKey'] = RAW_KEY
            res = requests.get(url, params=params, timeout=30).json()
            items = res['response']['body']['items'].get('item', [])
        return pd.DataFrame(items)
    except: return pd.DataFrame()

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    day_of_week = now.weekday() 

    df = pd.DataFrame()
    analysis_info = ""

    # 1. 월요일: 주간평균리포트 (기존 로직 유지)
    if day_of_week == 0: 
        mode_name = "주간평균리포트"
        weekly_dfs = []
        for i in range(3, 8): # 지난주 월~금
            target_dt = (now - datetime.timedelta(days=i)).strftime('%Y%m%d')
            df_day = get_official_data(target_dt)
            if not df_day.empty:
                for col in ['mkp', 'clpr', 'lopr', 'hipr', 'fltRt', 'trqu']:
                    df_day[col] = pd.to_numeric(df_day[col], errors='coerce').fillna(0)
                weekly_dfs.append(df_day)
        if weekly_dfs:
            full_df = pd.concat(weekly_dfs)
            df = full_df.groupby(['itmsNm', 'srtnCd', 'mrktCtg']).agg({
                'mkp': 'first', 'clpr': 'last', 'lopr': 'min', 'hipr': 'max', 'fltRt': 'mean', 'trqu': 'mean'
            }).reset_index()
            df.columns = ['종목명', '종목코드', '시장', '시가', '종가', '저가', '고가', '등락률(%)', '거래량']
            analysis_info = "기간: 지난주 월요일 ~ 금요일"

    # 2. 화~일: 일일마감리포트 (날짜 고정 로직 적용)
    else:
        mode_name = "일일마감리포트"
        # 어제 날짜로 고정 (주말 제외 로직은 공공데이터가 알아서 데이터 없음으로 반환)
        search_date = (now - datetime.timedelta(days=1)).strftime('%Y%m%d')
        print(f"📡 어제({search_date}) 데이터 조회 시도...")
        raw = get_official_data(search_date)
        
        # 데이터가 없으면 3월 3일로 넘어가지 않고 여기서 중단!
        if raw.empty or len(raw) < 500:
            async with bot:
                await bot.send_message(CHAT_ID, f"⚠️ 국장 알림: {search_date} 데이터가 공공데이터 서버에 아직 없습니다. (업데이트 대기 중)")
            return

        df = pd.DataFrame()
        df['종목코드'] = raw['srtnCd']
        df['종목명'] = raw['itmsNm']
        df['시가'] = pd.to_numeric(raw['mkp']).fillna(0).astype(int)
        df['종가'] = pd.to_numeric(raw['clpr']).fillna(0).astype(int)
        df['저가'] = pd.to_numeric(raw['lopr']).fillna(0).astype(int)
        df['고가'] = pd.to_numeric(raw['hipr']).fillna(0).astype(int)
        df['등락률(%)'] = pd.to_numeric(raw['fltRt']).fillna(0).astype(float)
        df['거래량'] = pd.to_numeric(raw['trqu']).fillna(0).astype(int)
        df['시장'] = raw['mrktCtg']
        analysis_info = f"마감 기준일: {search_date}"

    if df.empty: return

    # [이후 엑셀 생성 및 스타일링 - 지수님 요청 사항 100% 반영]
    target_cols = ['종목코드', '종목명', '시가', '종가', '저가', '고가', '등락률(%)', '거래량']
    file_name = f"{now.strftime('%m%d')}_{mode_name}.xlsx"
    
    header_fill = PatternFill(start_color="444444", end_color="444444", fill_type="solid")
    font_white_bold = Font(color="FFFFFF", bold=True)
    colors = {
        'red': PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid"),
        'orange': PatternFill(start_color="FFCC00", end_color="FFCC00", fill_type="solid"),
        'yellow': PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    }
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    center_align = Alignment(horizontal='center', vertical='center')

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for market in ['KOSPI', 'KOSDAQ']:
            for is_up in [True, False]:
                sheet_label = f"{market}_{'상승' if is_up else '하락'}"
                m_cond = df['시장'].str.contains(market)
                r_cond = (df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5)
                sub_df = df[m_cond & r_cond].copy().sort_values('등락률(%)', ascending=not is_up)
                sub_df = sub_df[target_cols]
                
                sub_df.to_excel(writer, sheet_name=sheet_label, index=False)
                ws = writer.sheets[sheet_label]

                ws.column_dimensions['A'].width = 12
                ws.column_dimensions['B'].width = 30
                ws.column_dimensions['C'].width = 15
                ws.column_dimensions['D'].width = 15
                ws.column_dimensions['E'].width = 15
                ws.column_dimensions['F'].width = 15
                ws.column_dimensions['G'].width = 12
                ws.column_dimensions['H'].width = 20

                for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=8), 1):
                    for c_idx, cell in enumerate(row, 1):
                        cell.border = border
                        cell.alignment = center_align
                        if r_idx == 1:
                            cell.fill, cell.font = header_fill, font_white_bold
                        else:
                            if c_idx in [3, 4, 5, 6, 8]: cell.number_format = '#,##0'
                            if c_idx == 7:
                                cell.number_format = '0.00'
                                val = abs(float(cell.value or 0))
                                name_cell = ws.cell(row=r_idx, column=2)
                                if val >= 25: 
                                    name_cell.fill, name_cell.font = colors['red'], font_white_bold
                                elif val >= 20: name_cell.fill = colors['orange']
                                elif val >= 10: name_cell.fill = colors['yellow']

    total_up = len(df[df['등락률(%)'] >= 5])
    total_down = len(df[df['등락률(%)'] <= -5])
    
    msg = (f"📊 모드: {mode_name}\n🔍 {analysis_info}\n\n"
           f"📈 상승(5%↑): {total_up}개\n📉 하락(5%↓): {total_down}개\n"
           f"💡 🟡10%↑ 🟠20%↑ 🔴25%↑")

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
