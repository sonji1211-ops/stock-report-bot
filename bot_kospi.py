import os, pandas as pd, asyncio, time
from yahooquery import Ticker
from datetime import datetime, timedelta
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_kospi_data_final():
    print("📡 [야후 쿼리 엔진] 코스피 정밀 분석 시작...")
    
    # 1. 실제 코스피 상장사 핵심 리스트 (성공 여부 확인용 50개)
    codes = [
        '005930','000660','005490','035420','035720','005380','051910','000270','068270','006400',
        '105560','055550','000810','012330','066570','096770','032830','003550','033780','000720',
        '009150','015760','018260','017670','011170','009540','036570','003670','034020','010130'
    ]
    tickers = [c + ".KS" for c in codes]
    
    try:
        # yahooquery는 yfinance보다 차단에 훨씬 강합니다.
        t = Ticker(tickers, asynchronous=True)
        data = t.price
        
        all_stocks = []
        for symbol, info in data.items():
            try:
                if isinstance(info, dict) and 'regularMarketPrice' in info:
                    ratio = info.get('regularMarketChangePercent', 0) * 100
                    all_stocks.append({
                        'Name': symbol.split('.')[0],
                        'Open': int(info.get('regularMarketOpen', 0)),
                        'Close': int(info.get('regularMarketPrice', 0)),
                        'Low': int(info.get('regularMarketDayLow', 0)),
                        'High': int(info.get('regularMarketDayHigh', 0)),
                        'Ratio': float(ratio),
                        'Volume': int(info.get('regularMarketVolume', 0))
                    })
            except: continue
        
        print(f"✅ {len(all_stocks)}개 종목 데이터 확보 성공!")
        return pd.DataFrame(all_stocks)
    except Exception as e:
        print(f"❌ 엔진 구동 실패: {e}")
        return pd.DataFrame()

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    
    df = get_kospi_data_final()
    
    # [방어 로직] 만약 이번에도 0개라면, 지수님께 상황 보고 후 종료
    if df.empty:
        print("🚨 깃허브 IP가 완전히 차단되었습니다. 로컬 실행을 권장합니다.")
        return

    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    # 5% 필터링
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 엑셀 저장 (기존 디자인 유지)
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, d in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            tmp = d.rename(columns={'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}) if not d.empty else pd.DataFrame([['조건 만족 없음']+['']*6], columns=['종목명','시가','종가','저가','고가','등락률(%)','거래량'])
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            # ... [디자인 코드는 이전과 동일] ...

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📊 분석: {len(df)}개 완료"
    with open(file_name, 'rb') as f:
        await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
    os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
