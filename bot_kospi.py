import os, pandas as pd, requests, re, io, time, random, asyncio
from datetime import datetime, timedelta
from telegram import Bot
from bs4 import BeautifulSoup
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

TOKEN = "8574978661:AAF5SXIgfpJlnAfN5ccSk0tJek_uSlCMBBo"
CHAT_ID = "8564327930"

def fetch_naver_stock(sosok, page):
    # 세션 사용으로 쿠키 유지 (차단 방지 핵심)
    session = requests.Session()
    url = f"https://finance.naver.com/sise/sise_market_sum.naver?sosok={sosok}&field=quant&field=open&field=high&field=low&field=frate&page={page}"
    
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
        'Referer': 'https://finance.naver.com/sise/',
        'Accept-Language': 'ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7'
    }
    
    try:
        resp = session.get(url, headers=headers, timeout=15)
        # 한글 깨짐 방지
        resp.encoding = 'euc-kr' 
        
        if resp.status_code != 200:
            print(f"❌ 접속 실패 (상태 코드: {resp.status_code})")
            return []
            
        soup = BeautifulSoup(resp.text, 'html.parser')
        # 네이버 금융 특유의 테이블 구조 타겟팅
        table = soup.select_one('table.type_2')
        
        if not table:
            print(f"⚠️ {page}p: 테이블을 찾을 수 없습니다. (차단 가능성 높음)")
            return []
        
        rows = []
        # 데이터가 있는 tr만 추출 (선이 있는 줄 제외)
        tr_list = table.select('tr')
        
        for tr in tr_list:
            tds = tr.select('td')
            if len(tds) < 10: continue
            
            # 종목명 추출 (a 태그 우선)
            name_tag = tds[1].select_one('a')
            if not name_tag: continue
            name = name_tag.get_text(strip=True)
            
            def clean(i):
                val = tds[i].get_text(strip=True).replace(',', '').replace('%', '').replace('+', '').replace('-', '0')
                return val if val else '0'
                
            try:
                rows.append({
                    'Name': name,
                    'Close': int(clean(2)),
                    'Ratio': float(tds[4].get_text(strip=True).replace('%','').replace('+','')),
                    'Volume': int(clean(5)),
                    'Open': int(clean(7)),
                    'High': int(clean(8)),
                    'Low': int(clean(9))
                })
            except Exception as e:
                continue
                
        return rows
    except Exception as e:
        print(f"⚠️ {page}p 에러: {e}")
        return []

async def run_report():
    bot = Bot(token=TOKEN)
    now = datetime.utcnow() + timedelta(hours=9)
    all_data = []
    
    print(f"📡 KOSPI 전수조사 시작 ({now.strftime('%Y-%m-%d %H:%M')})")
    
    for p in range(1, 31):
        data = fetch_naver_stock(0, p)
        if not data:
            # 1페이지에서 실패하면 한 번 더 시도 (랜덤 대기 후)
            if p == 1:
                print("🔄 1페이지 재시도 중...")
                time.sleep(3)
                data = fetch_naver_stock(0, p)
            if not data: break
            
        all_data.extend(data)
        print(f"✅ {p}/30p 수집 완료 (현재 {len(all_data)}개)")
        # 네이버 감시를 피하기 위한 랜덤 지연
        time.sleep(random.uniform(1.0, 2.5))

    if not all_data:
        print("❌ 수집된 데이터가 없습니다. 프로그램을 종료합니다.")
        return

    df = pd.DataFrame(all_data)
    r_type = "주간평균" if now.weekday() == 6 else "일일"
    file_name = f"{now.strftime('%m%d')}_KOSPI_{r_type}.xlsx"
    
    # 지수님 요구사항 필터링 (5% 이상/이하)
    up_df = df[df['Ratio'] >= 5.0].sort_values('Ratio', ascending=False)
    down_df = df[df['Ratio'] <= -5.0].sort_values('Ratio', ascending=True)

    # 엑셀 디자인 및 포맷 (기존 요구사항 100% 유지)
    h_map = {'Name':'종목명','Open':'시가','Close':'종가','Low':'저가','High':'고가','Ratio':'등락률(%)','Volume':'거래량'}
    red, ora, yel = PatternFill("solid", "FF0000"), PatternFill("solid", "FFCC00"), PatternFill("solid", "FFFF00")
    header_f, white_f = PatternFill("solid", "444444"), Font(color="FFFFFF", bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        for s_name, d in {'코스피_상승': up_df, '코스피_하락': down_df}.items():
            tmp = d.rename(columns=h_map) if not d.empty else pd.DataFrame([['조건 만족 종목 없음']+['']*6], columns=list(h_map.values()))
            tmp.to_excel(writer, sheet_name=s_name, index=False)
            ws = writer.sheets[s_name]
            for cell in ws[1]: cell.fill, cell.font, cell.alignment = header_f, white_f, Alignment(horizontal='center')
            for r in range(2, ws.max_row + 1):
                try:
                    val = ws.cell(r, 6).value
                    v = abs(float(val)) if val and str(val).replace('.','').replace('-','').isdigit() else 0
                    if v >= 28: ws.cell(r, 1).fill, ws.cell(r, 1).font = red, white_f
                    elif v >= 20: ws.cell(r, 1).fill = ora
                    elif v >= 10: ws.cell(r, 1).fill = yel
                except: pass
                for c in range(1, 8):
                    ws.cell(r, c).alignment, ws.cell(r, c).border = Alignment(horizontal='center'), border
                    if c in [2,3,4,5,7]: ws.cell(r, c).number_format = '#,##0'
                    if c == 6: ws.cell(r, c).number_format = '0.00'
            ws.column_dimensions['A'].width = 18

    msg = f"📅 {now.strftime('%m-%d')} *[KOSPI {r_type}]*\n📊 전체 수집: {len(df)}개\n📈 상승(5%↑): {len(up_df)} / 📉 하락(5%↓): {len(down_df)}"
    try:
        with open(file_name, 'rb') as f:
            await bot.send_document(CHAT_ID, document=f, caption=msg, parse_mode="Markdown")
        print("🚀 텔레그램 리포트 발송 성공!")
    except Exception as e:
        print(f"❌ 발송 에러: {e}")
    finally:
        if os.path.exists(file_name): os.remove(file_name)

if __name__ == "__main__":
    asyncio.run(run_report())
