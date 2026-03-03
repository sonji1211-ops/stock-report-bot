import os, pandas as pd, asyncio, time
from yahooquery import Ticker
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_kospi_data_final():
    print("📡 [야후 쿼리 엔진] 코스피 전수조사 가동...")
    
    # [종목 확장] 지수님이 원하시는 넉넉한 분석을 위해 코스피 주요 종목 리스트 확보
    # 시총 상위 및 변동성이 큰 핵심 종목 위주
    codes = [
        '005930','000660','005490','035420','035720','005380','051910','000270','068270','006400',
        '105560','055550','000810','012330','066570','096770','032830','003550','033780','000720',
        '009150','015760','018260','017670','011170','009540','036570','003670','034020','010130',
        '010950','251270','000100','008930','086790','004020','078930','028260','000120','030200',
        '039130','011070','000080','005070','009830','001570','016360','004170','036460','010120',
        '010140','001450','003410','000060','000210','001040','001740','002350','002790','003490',
        '003520','004370','004800','004990','005830','006120','006260','006360','007070','007310',
        '008770','009240','010060','010620','011780','011790','012450','014680','014820','017800'
    ]
    tickers = [c + ".KS" for c in codes]
    
    try:
        t = Ticker(tickers, asynchronous=True)
        data = t.price
        
        all_stocks = []
        for symbol, info in data.items():
            if isinstance(info, dict) and 'regularMarketPrice' in info:
                # 등락률 계산 (소수점 유지)
                ratio = info.get('regularMarketChangePercent', 0) * 100
                all_stocks.append({
                    'Name': symbol.split('.')[0],
                    'Open': info.get('regularMarketOpen', 0),
                    'Close': info.get('regularMarketPrice', 0),
                    'Low': info.get('regularMarketDayLow', 0),
                    'High': info.get('regularMarketDayHigh', 0),
                    'Ratio': float(ratio),
                    'Volume': info.get('regularMarketVolume', 0)
                })
        
        print(f"✅ {len(all_stocks)}개 종목 데이터 확보 성공!")
        return pd.DataFrame(all_stocks)
    except Exception as e:
        print(f"❌ 엔진 구동 실패: {e}")
        return pd.DataFrame()

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    df = get_kospi_data_final()
    
    if df.empty:
        print("🚨 데이터가 비어있습니다. 전송을 중단합니다.")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    # [필터링] 지수님 요구사항: 등락률 5% 이상/이하
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # [디자인 세팅]
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red_f = PatternFill("solid", "FF0000")   # 28% 이상
    ora_f = PatternFill("solid", "FFCC00")   # 20% 이상
    yel_f = PatternFill("solid", "FFFF00")   # 10% 이상
    head_f = PatternFill("solid", "444444")
    white_font = Font(color="FFFFFF", bold=True)
    center_align = Alignment(horizontal='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, d in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            # 데이터가 없을 때 처리
            if d.empty:
                tmp = pd.DataFrame([['조건 만족 종목 없음']+['']*6], columns=list(h_map.values()))
            else:
                tmp = d.rename(columns=h_map)
            
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            
            # 헤더 스타일
            for cell in ws[1]:
                cell.fill = head_f
                cell.font = white_font
                cell.alignment = center_align
                cell.border = thin_border

            # 본문 스타일 및 조건부 서식
            for r in range(2, ws.max_row + 1):
                # 등락률 값 기준 색상 (6번째 열)
                try:
                    ratio_val = abs(float(ws.cell(r, 6).value))
                    if ratio_val >= 28:
                        ws.cell(r, 1).fill = red_f
                        ws.cell(r, 1).font = white_font
                    elif ratio_val >= 20:
                        ws.cell(r, 1).fill = ora_f
                    elif ratio_val >= 10:
                        ws.cell(r, 1).fill = yel_f
                except: pass

                for c in range(1, 8):
                    ws.cell(r, c).alignment = center_align
                    ws.cell(r, c).border = thin_border
                    # 숫자 포맷 (시가, 종가, 저가, 고가, 거래량은 콤마)
                    if c in [2, 3, 4, 5, 7]:
                        ws.cell(r, c).number_format = '#,##0'
                    # 등락률은 소수점 2자리
                    if c == 6:
                        ws.cell(r, c).number_format = '0.00'
            
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['F'].width = 12
            ws.column_dimensions['G'].width = 15

    # 텔레그램 전송
    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📈 분석대상: {len(df)}개\n🚀 상승(5%↑): {len(up_df)}개 / 하락(5%↓): {len(down_df)}개"
    try:
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        print("🚀 전송 성공!")
    except Exception as e:
        print(f"❌ 전송 실패: {e}")
    finally:
        if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
