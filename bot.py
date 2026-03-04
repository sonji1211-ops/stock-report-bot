import os, pandas as pd, asyncio, datetime, requests, time
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
from urllib.parse import unquote

# [1. 설정 정보]
TELEGRAM_TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"
RAW_KEY = "3e937f2b0780c88e27c6f4cb99d5b58e69cc71cef898809e7aacb2bcabe1b438"
SERVICE_KEY = unquote(RAW_KEY)

def get_stock_data(target_date):
    """특정 날짜의 주식 데이터를 가져오는 함수"""
    url = 'http://apis.data.go.kr/1160100/service/GetStockSecuritiesInfoService/getStockPriceInfo'
    params = {'serviceKey': SERVICE_KEY, 'numOfRows': '4000', 'resultType': 'json', 'basDt': target_date}
    try:
        response = requests.get(url, params=params, timeout=30)
        res = response.json()['response']['body']['items'].get('item', [])
        if not res:
            params['serviceKey'] = RAW_KEY
            response = requests.get(url, params=params, timeout=30)
            res = response.json()['response']['body']['items'].get('item', [])
        return pd.DataFrame(res)
    except:
        return pd.DataFrame()

async def main():
    bot = Bot(token=TELEGRAM_TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    day_of_week = now.weekday() # 0:월, 6:일

    # [2. 주간/일일 데이터 수집 및 처리]
    if day_of_week == 6: # 일요일 (주간 평균 리포트)
        mode = "주간평균리포트"
        weekly_dfs = []
        print("LOG: 주간 데이터(월~금) 수집 시작...")
        for i in range(2, 7): # 일요일 기준 2일전(금) ~ 6일전(월)
            target_dt = (now - datetime.timedelta(days=i)).strftime('%Y%m%d')
            df_day = get_stock_data(target_dt)
            if not df_day.empty:
                # 계산을 위해 숫자형 변환
                df_day['clpr'] = pd.to_numeric(df_day['clpr'], errors='coerce').fillna(0)
                df_day['fltRt'] = pd.to_numeric(df_day['fltRt'], errors='coerce').fillna(0)
                df_day['trqu'] = pd.to_numeric(df_day['trqu'], errors='coerce').fillna(0)
                weekly_dfs.append(df_day)
            time.sleep(0.3)
        
        if not weekly_dfs:
            print("LOG: 수집된 주간 데이터가 없습니다."); return
        
        full_df = pd.concat(weekly_dfs)
        # 종목별 평균 계산
        df = full_df.groupby(['itmsNm', 'mrktCtg', 'srtnCd']).agg({
            'clpr': 'mean', 'fltRt': 'mean', 'trqu': 'mean'
        }).reset_index()
        df.columns = ['종목명', '시장', '종목코드', '종가(평균)', '등락률(%)', '거래량(평균)']
    else: # 평일 (일일 마감 리포트)
        mode = "일일마감"
        print(f"LOG: {now.strftime('%Y%m%d')} 데이터 수집 시작...")
        raw = get_stock_data(now.strftime('%Y%m%d'))
        if raw.empty:
            print("LOG: 오늘자 데이터가 아직 공시되지 않았습니다."); return
        
        df = pd.DataFrame()
        df['시장'] = raw['mrktCtg']
        df['종목코드'] = raw['srtnCd']
        df['종목명'] = raw['itmsNm']
        df['종가'] = pd.to_numeric(raw['clpr']).fillna(0).astype(int)
        df['등락률(%)'] = pd.to_numeric(raw['fltRt']).fillna(0).astype(float)
        df['거래량'] = pd.to_numeric(raw['trqu']).fillna(0).astype(int)

    # [3. 엑셀 파일 생성 및 디자인]
    file_name = f"{now.strftime('%m%d_%H%M')}_{mode}.xlsx"
    h_fill = PatternFill("solid", fgColor="444444")
    f_white_b = Font(color="FFFFFF", bold=True)
    colors = {'red': PatternFill("solid", "FF0000"), 'orange': PatternFill("solid", "FFCC00"), 'yellow': PatternFill("solid", "FFFF00")}
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for m_name in ['KOSPI', 'KOSDAQ']:
            for is_up in [True, False]:
                # ±5% 이상 필터링
                cond = (df['시장'] == m_name) & ((df['등락률(%)'] >= 5) if is_up else (df['등락률(%)'] <= -5))
                sub_df = df[cond].sort_values('등락률(%)', ascending=not is_up).drop(columns=['시장'])
                
                sheet_label = f"{m_name}_{'상승' if is_up else '하락'}"
                sub_df.to_excel(writer, sheet_name=sheet_label, index=False)
                
                ws = writer.sheets[sheet_label]
                # 컬럼 너비 설정 (A: 종목명 30, 마지막: 거래량 25)
                ws.column_dimensions['A'].width = 30
                ws.column_dimensions['B'].width = 15
                ws.column_dimensions['C'].width = 15
                ws.column_dimensions['D'].width = 15
                ws.column_dimensions['E'].width = 25 # 거래량

                for r in range(1, ws.max_row + 1):
                    for c in range(1, ws.max_column + 1):
                        cell = ws.cell(r, c)
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                        cell.border = border
                        
                        if r == 1: # 헤더 디자인
                            cell.fill, cell.font = h_fill, f_white_b
                        else:
                            # 숫자 포맷팅
                            if "종가" in ws.cell(1, c).value or "거래량" in ws.cell(1, c).value:
                                cell.number_format = '#,##0'
                            if "등락률" in ws.cell(1, c).value:
                                cell.number_format = '0.00'
                                val = abs(float(cell.value or 0))
                                target = ws.cell(r, 1) # 종목명 강조
                                if val >= 25: 
                                    target.fill, target.font = colors['red'], f_white_b
                                elif val >= 20: 
                                    target.fill = colors['orange']
                                elif val >= 10: 
                                    target.fill = colors['yellow']

    # [4. 텔레그램 메시지 및 전송]
    up_cnt = len(df[df['등락률(%)'] >= 5])
    down_cnt = len(df[df['등락률(%)'] <= -5])
    msg = (f"📅 {now.strftime('%Y-%m-%d')} {mode}\n"
           f"━━━━━━━━━━━━━━━━━━\n"
           f"📈 상승(5%↑): {up_cnt}개 종목\n"
           f"📉 하락(5%↓): {down_cnt}개 종목\n"
           f"━━━━━━━━━━━━━━━━━━\n"
           f"✅ 가독성: 거래량 열 너비 25 적용\n"
           f"✅ 주간: 월~금 5일 평균 등락률 기준")

    async with bot:
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=msg)
    if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
