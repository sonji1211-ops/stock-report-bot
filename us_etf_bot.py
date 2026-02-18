import os
import FinanceDataReader as fdr
import pandas as pd
from datetime import datetime, timedelta
import asyncio
from telegram import Bot
from openpyxl.styles import Alignment, PatternFill, Font

# [설정] 텔레그램 정보
TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

# [통합 자산 리스트] 누락 없는 40종 ETF + 국장 + 원화 코인
ASSET_NAMES = {
    # 1. 국내 지수 및 환율
    'KS11': '코스피 지수', 'KQ11': '코스닥 지수', 'USD/KRW': '달러/원 환율',
    
    # 2. 국내 주요 ETF (455810 포함 안정화)
    '069500': 'KODEX 200', '252670': 'KODEX 200선물인버스2X', '305720': 'KODEX 2차전지산업',
    '455810': 'TIGER 미국배당다우존스', '462330': 'KODEX AI반도체핵심공정', '122630': 'KODEX 레버리지',
    
    # 3. 가상화폐 (가장 안정적인 원화 데이터 포맷: -KRW)
    'BTC-KRW': '비트코인', 'ETH-KRW': '이더리움', 'XRP-KRW': '리플(XRP)', 
    'SOL-KRW': '솔라나(SOL)', 'USDT-KRW': '테더(USDT)',
    
    # 4. 미국 지수 및 주요 ETF (주셨던 40종 전체 복구)
    'QQQ': '나스닥100', 'TQQQ': '나스닥100(3배)', 'SQQQ': '나스닥100인버스(3배)', 'QLD': '나스닥100(2배)',
    'SPY': 'S&P500', 'IVV': 'S&P500(iShares)', 'VOO': 'S&P500(Vanguard)', 'SSO': 'S&P500(2배)', 'Upro': 'S&P500(3배)',
    'DIA': '다우존스', 'IWM': '러셀2000', 
    'SOXX': '필라델피아반도체', 'SOXL': '반도체강세(3배)', 'SOXS': '반도체약세(3배)', 'SMH': '반도체ETF(VanEck)',
    'NVDL': '엔비디아(2배)', 'TSLL': '테슬라(2배)', 'CONL': '코인베이스(2배)',
    'SCHD': '슈드(배당성장)', 'JEPI': '제피(고배당)', 'ARKK': '아크혁신(캐시우드)',
    'TLT': '미국채20년(장기채)', 'TMF': '장기채강세(3배)', 'TMV': '장기채약세(3배)',
    'XLF': '금융섹터', 'XLV': '헬스케어섹터', 'XLE': '에너지섹터', 'XLK': '기술주섹터', 
    'XLY': '임의소비재', 'XLP': '필수소비재', 'GDX': '금광업', 'GLD': '금선물',
    'VNQ': '리츠(부동산)', 'BITO': '비트코인ETF', 'FNGU': '빅테크플러스(3배)', 'BULZ': '빅테크성장(3배)',
    'VTI': '미국전체주식', 'VXUS': '미국외전세계', 'VT': '전세계주식',
    
    # 5. 원자재
    'GC=F': '금 선물', 'SI=F': '은 선물'
}

async def fetch_asset_data(symbol, search_start, search_end, mode):
    try:
        # 데이터 소스에 따른 수집 로직 보강
        df = fdr.DataReader(symbol, search_start, search_end)
        
        if df is None or df.empty or len(df) < 2:
            return None
        
        # 마지막 유효 데이터 2개 행 추출 (등락률 계산 핵심)
        last_val = df.iloc[-1]
        prev_val = df.iloc[-2]
        
        last_close = float(last_val['Close'])
        prev_close = float(prev_val['Close'])
        
        if mode == 'daily':
            ratio = round(((last_close - prev_close) / prev_close) * 100, 2)
            final_date = df.index[-1].strftime('%Y-%m-%d')
        else:
            first_open = float(df.iloc[0]['Open'])
            ratio = round(((last_close - first_open) / first_open) * 100, 2)
            final_date = f"{df.index[0].strftime('%m%d')}~{df.index[-1].strftime('%m%d')}"
            
        return {'티커': symbol, '항목명': ASSET_NAMES.get(symbol, symbol), '현재가': last_close, '등락률': ratio, '기준일': final_date}
    except Exception:
        return None

async def send_etf_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    day_of_week = now.weekday()
    
    # 코인 주말 데이터를 위해 20일 전부터 데이터 확보
    search_end = now.strftime('%Y-%m-%d')
    search_start = (now - timedelta(days=20)).strftime('%Y-%m-%d')
    mode = 'weekly' if day_of_week == 6 else 'daily'

    tasks = [fetch_asset_data(s, search_start, search_end, mode) for s in ASSET_NAMES.keys()]
    results = await asyncio.gather(*tasks)
    df_raw = pd.DataFrame([r for r in results if r is not None])
    
    if df_raw.empty: return

    most_common_date = df_raw['기준일'].value_counts().idxmax()
    df_final = df_raw.copy()

    file_name = f"{now.strftime('%m%d')}_종합_자산_리포트.xlsx"
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        df_final[['티커','항목명','현재가','등락률']].rename(columns={'등락률':'등락률(%)'}).to_excel(writer, sheet_name='종합현황', index=False)
        ws = writer.sheets['종합현황']
        
        # 가독성을 위한 너비 조정
        ws.column_dimensions['A'].width = 15
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 22
        ws.column_dimensions['D'].width = 15
        
        for row in range(1, ws.max_row + 1):
            for col in range(1, 5):
                cell = ws.cell(row, col)
                if col == 2:
                    cell.alignment = Alignment(horizontal='left', vertical='center', indent=1)
                else: # 티커, 현재가, 등락률은 중앙 정렬
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                if row > 1:
                    ticker_str = str(ws.cell(row, 1).value)
                    # 원화 자산(코인, 국주)은 정수형, 달러 자산은 소수점 2자리
                    if '-KRW' in ticker_str or ticker_str.isdigit() or ticker_str in ['KS11', 'KQ11']:
                        ws.cell(row, 3).number_format = '#,##0'
                    else:
                        ws.cell(row, 3).number_format = '#,##0.00'
                    
                    ws.cell(row, 4).number_format = '0.00'
                    
                    # 3% 이상 변동 시 강조
                    ratio_val = float(ws.cell(row, 4).value or 0)
                    if col == 2 and abs(ratio_val) >= 3:
                        cell.fill = PatternFill("solid", fgColor="FFFF00")
                        cell.font = Font(bold=True)

    async with bot:
        title = "🌍 [종합]" if mode == 'daily' else "🗓 [주간]"
        await bot.send_document(CHAT_ID, open(file_name, 'rb'), caption=f"{title} 한·미 자산 리포트 ({most_common_date})\n💡 455810 에러 및 코인 등락률 누락 완전 수정 완료")

if __name__ == "__main__":
    asyncio.run(send_etf_report())