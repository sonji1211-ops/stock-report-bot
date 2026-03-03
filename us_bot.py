import pandas as pd
import yfinance as yf
import datetime, os, asyncio
from telegram import Bot

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def get_real_tickers():
    """실제로 상장사가 밀집된 구간만 골라서 티커 리스트 생성"""
    # 000XXX 대역은 없는 번호가 너무 많아 404 에러의 주범입니다.
    # 실속 있는 구간 (삼성전자 005930, 현대차 005380 등) 위주로 구성
    ranges = [
        (10, 2000, 10),     # 000010 ~ 002000
        (5000, 15000, 50),  # 005000 ~ 015000 (우량주 밀집)
        (20000, 40000, 100),# 020000 ~ 040000
        (50000, 150000, 200)# 코스닥 중대형주 구간
    ]
    
    tickers = []
    for start, end, step in ranges:
        for i in range(start, end, step):
            code = f"{i:06d}"
            tickers.append(f"{code}.KS")
            tickers.append(f"{code}.KQ")
    return tickers

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    tickers = get_real_tickers()
    print(f"📡 [1단계] {len(tickers)}개 후보 중 유효 종목 수집 시작...")

    # threads=True와 progress=False로 속도와 로그를 깔끔하게 유지
    # 50개씩 끊어서 야후 차단을 회피합니다.
    all_data = []
    chunk_size = 50
    
    for i in range(0, len(tickers), chunk_size):
        batch = tickers[i:i+chunk_size]
        try:
            # 2일치 데이터를 가져와서 전일 대비 등락률 확보
            df = yf.download(batch, period="2d", interval="1d", group_by='ticker', threads=True, progress=False)
            
            for t in batch:
                if t not in df.columns.levels[0]: continue
                stock = df[t].dropna()
                if len(stock) < 2: continue
                
                curr_c = stock['Close'].iloc[-1]
                prev_c = stock['Close'].iloc[-2]
                vol = stock['Volume'].iloc[-1]
                
                # 시세가 있고 거래량이 0이 아닌 경우만 저장
                if curr_c > 0 and vol > 0:
                    all_data.append({
                        'Code': t.split('.')[0],
                        'Market': "KOSPI" if t.endswith(".KS") else "KOSDAQ",
                        'Close': int(curr_c),
                        'Ratio': float(((curr_c - prev_c) / prev_c) * 100),
                        'Volume': int(vol)
                    })
        except:
            continue
        print(f"📦 {i+chunk_size}개 처리 중... 현재 확보: {len(all_data)}개")

    final_df = pd.DataFrame(all_data)
    print(f"✅ 최종 수집 완료: {len(final_df)}개 유효 종목 확보")

    # [임시 파일 생성 및 전송] - 데이터 확인용
    if not final_df.empty:
        file_name = "test_data.xlsx"
        final_df.to_excel(file_name, index=False)
        
        async with bot:
            msg = (f"📅 {now.strftime('%Y-%m-%d')} 수집 결과\n\n"
                   f"📊 유효 종목: {len(final_df)}개 확보\n"
                   f"🚀 수집 우선순위 테스트 완료")
            with open(file_name, 'rb') as f:
                await bot.send_document(CHAT_ID, f, caption=msg)
        os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
