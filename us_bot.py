import pandas as pd
import yfinance as yf
import FinanceDataReader as fdr
import datetime, os, asyncio, time
from telegram import Bot

# [설정]
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

async def main():
    bot = Bot(token=TOKEN)
    now = datetime.datetime.utcnow() + datetime.timedelta(hours=9)
    
    print("📡 [1단계] KRX/KOSDAQ 전 종목 리스트 확보 중...")
    try:
        # 난수 대신 실제 상장된 종목(약 2,700개) 리스트를 가져옵니다.
        df_krx = fdr.StockListing('KRX')
        # 야후용 티커 변환 (6자리 코드 + .KS 또는 .KQ)
        df_krx['Ticker'] = df_krx['Code'].apply(lambda x: x + (".KS" if x.isdigit() and int(x) < 900000 else ".KQ"))
        all_tickers = df_krx['Ticker'].tolist()
    except Exception as e:
        print(f"❌ 종목 리스트 확보 실패: {e}")
        return

    print(f"🚀 총 {len(all_tickers)}개 실제 종목 분석 시작 (거래량 상위 800개 타겟)...")

    # 1. 데이터 수집 (100개씩 묶어서 야후에 요청)
    collected_data = []
    chunk_size = 100
    
    for i in range(0, len(all_tickers), chunk_size):
        batch = all_tickers[i:i+chunk_size]
        try:
            # 2일치 데이터를 가져와야 등락률 계산이 가능합니다.
            data = yf.download(batch, period="2d", interval="1d", group_by='ticker', threads=True, progress=False)
            
            for t in batch:
                if t not in data.columns.levels[0]: continue
                df_t = data[t].dropna()
                if len(df_t) < 2: continue
                
                vol = df_t['Volume'].iloc[-1]
                curr_c = df_t['Close'].iloc[-1]
                prev_c = df_t['Close'].iloc[-2]
                
                if vol > 100: # 거래가 실질적으로 발생하는 종목만
                    ratio = ((curr_c - prev_c) / prev_c) * 100
                    collected_data.append({
                        'Code': t.split('.')[0],
                        'Market': "KOSPI" if t.endswith(".KS") else "KOSDAQ",
                        'Close': int(curr_c),
                        'Ratio': float(ratio),
                        'Volume': int(vol)
                    })
        except: continue
        print(f"📦 {min(i+chunk_size, len(all_tickers))}개 스캔 완료... 현재 확보: {len(collected_data)}개")

    if not collected_data:
        print("🚨 유효 데이터를 수집하지 못했습니다.")
        return

    # 2. 거래량 기준 상위 800개 추출 (지수님 요청 1순위)
    full_df = pd.DataFrame(collected_data)
    top_800_vol = full_df.sort_values(by='Volume', ascending=False).head(800)
    
    # 3. 그 800개 중 등락률 ±5% 이상만 필터링 (지수님 요청 2순위)
    final_filtered = top_800_vol[(top_800_vol['Ratio'] >= 5) | (top_800_vol['Ratio'] <= -5)]
    
    # 4. 등락률 순 정렬
    final_df = final_filtered.sort_values(by='Ratio', ascending=False)

    print(f"✅ 필터링 완료: 최종 {len(final_df)}개 종목 확보")

    # [데이터 확인용 엑셀 전송]
    file_name = "data_check.xlsx"
    final_df.to_excel(file_name, index=False)
    
    async with bot:
        msg = (f"📅 {now.strftime('%Y-%m-%d')} 데이터 우선 수집 결과\n\n"
               f"📊 실제 거래 상위 800개 분석\n"
               f"⚡ 등락률 ±5% 필터 적용 완료\n"
               f"📈 수집된 유효 종목: {len(final_df)}개")
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, f, caption=msg)
    os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(main())
