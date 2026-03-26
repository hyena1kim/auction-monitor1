import pandas as pd
import time
import asyncio
import sys
import urllib.parse
import os

from playwright.async_api import async_playwright

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --- Streamlit 환경(배포)에서 드라이버 재생성 방지용: 있으면 사용, 없으면 무시 ---
try:
    import streamlit as st
except Exception:
    st = None


# =========================
# 0) 유틸: 시스템 바이너리 찾기
# =========================
def _first_existing_path(paths):
    for p in paths:
        if p and os.path.exists(p):
            return p
    return None


# =========================
# 1) Selenium 드라이버 생성 (Streamlit Cloud 안정화 핵심)
#    - Selenium Manager 캐시(/home/appuser/.cache/selenium/...)를 타지 않도록
#      /usr/bin/chromium + /usr/bin/chromedriver를 명시적으로 사용
# =========================
def _build_chrome_options() -> Options:
    options = Options()

    # 컨테이너/리눅스 환경에서는 headless가 사실상 필수
    options.add_argument("--headless=new")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")

    # 자동화 탐지 최소화(선택)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    return options


def make_driver() -> webdriver.Chrome:
    """
    Streamlit Cloud에서 가장 예측 가능한 구성:
    - OS(apt)로 설치된 chromium/chromedriver를 사용하도록 경로를 명시
    - Selenium Manager가 다운받는 캐시 chromedriver를 사용하지 않게 함
    """
    options = _build_chrome_options()

    # Streamlit Cloud(데비안)에서 흔한 chromium 경로들
    chromium_bin = _first_existing_path([
        "/usr/bin/chromium",
        "/usr/bin/chromium-browser",
        "/usr/bin/google-chrome",
        "/usr/bin/google-chrome-stable",
    ])
    if chromium_bin:
        options.binary_location = chromium_bin

    # chromedriver도 시스템 경로를 명시 (캐시 경로 사용 방지)
    chromedriver_bin = _first_existing_path([
        "/usr/bin/chromedriver",
        "/usr/lib/chromium/chromedriver",
        "/usr/lib/chromium-browser/chromedriver",
    ])
    if not chromedriver_bin:
        # 마지막 fallback: 그냥 Service()를 쓰면 Selenium Manager가 캐시로 받을 수 있음
        # 하지만 여기까지 오면 packages.txt 설치가 안 됐을 가능성이 큼
        raise RuntimeError(
            "chromedriver를 시스템에서 찾지 못했습니다. "
            "packages.txt에 chromium/chromium-driver가 설치되었는지 확인하세요."
        )

    service = Service(chromedriver_bin)
    driver = webdriver.Chrome(service=service, options=options)
    return driver


# Streamlit이면 드라이버를 캐시해 재실행시 반복 생성 방지 (선택)
if st is not None:
    @st.cache_resource(show_spinner=False)
    def get_driver_cached():
        return make_driver()
else:
    def get_driver_cached():
        return make_driver()


# =========================
# 2) 사이트별 스크래핑 함수들
# =========================
def scrape_seoul_auction(driver):
    print("\n--- 서울옥션 경매 정보 수집 시작 ---")
    url = "https://www.seoulauction.com/auction-list/upcoming"
    driver.get(url)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "auction_info"))
        )
    except Exception as e:
        print("서울옥션 데이터를 찾을 수 없거나 로딩 시간이 초과되었습니다.", e)
        return []

    time.sleep(2)
    auction_infos = driver.find_elements(By.CLASS_NAME, "auction_info")
    print(f"총 {len(auction_infos)}개의 서울옥션 경매 정보를 찾았습니다.")

    data_list = []
    for info in auction_infos:
        item = {}
        try:
            type_elements = info.find_elements(By.CLASS_NAME, "type")
            item["유형/상태"] = ", ".join([t.text for t in type_elements if t.text.strip()])
        except Exception:
            item["유형/상태"] = ""

        try:
            title_element = info.find_element(By.CLASS_NAME, "title")
            item["경매명"] = title_element.text.strip()
        except Exception:
            item["경매명"] = ""

        try:
            desc_area = info.find_element(By.CLASS_NAME, "description")
            dls = desc_area.find_elements(By.TAG_NAME, "dl")
            for dl in dls:
                dt_text = dl.find_element(By.TAG_NAME, "dt").text.strip()
                dd_text = dl.find_element(By.TAG_NAME, "dd").text.strip()
                item[dt_text] = dd_text
        except Exception:
            pass

        item["바로가기 URL"] = f'=HYPERLINK("{url}", "{url}")'
        data_list.append(item)
        print(f"서울옥션 추출 완료: {item.get('경매명', '이름 없음')}")

    return data_list


def scrape_k_auction(driver):
    print("\n--- 케이옥션 출품작 정보 수집 시작 ---")
    url = (
        "https://www.k-auction.com/Auction/Premium/225?"
        "page_size=10&page_type=P&auc_kind=2&auc_num=225&work_type=2672"
    )
    driver.get(url)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CLASS_NAME, "artwork"))
        )
    except Exception as e:
        print("케이옥션 데이터를 찾을 수 없거나 로딩 시간이 초과되었습니다.", e)
        return []

    auction_title = ""
    auction_schedule = ""
    auction_location = ""

    try:
        subtop = driver.find_element(By.CLASS_NAME, "subtop-desc")
        auction_title = subtop.find_element(By.TAG_NAME, "h1").text.strip()

        p_tag = subtop.find_element(By.TAG_NAME, "p")
        auction_schedule = p_tag.find_element(By.TAG_NAME, "strong").text.strip()
        auction_location = p_tag.find_element(By.TAG_NAME, "span").text.strip()
        print(f"케이옥션 정보 파싱 완료: {auction_title}")
    except Exception as e:
        print("케이옥션 상단 정보를 가져오는 중 에러 발생:", e)

    # 전체 페이지 로딩을 위해 스크롤 다운
    last_height = driver.execute_script("return document.body.scrollHeight")
    while True:
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(2)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height

    time.sleep(2)
    artworks = driver.find_elements(By.CLASS_NAME, "artwork")
    print(f"총 {len(artworks)}개의 케이옥션 출품작 정보를 찾았습니다.")

    data_list = []
    for art in artworks:
        item = {}

        try:
            item["Lot"] = art.find_element(By.CLASS_NAME, "lot").text.strip()
        except Exception:
            item["Lot"] = ""

        try:
            img_tag = art.find_element(By.TAG_NAME, "img")
            img_url = img_tag.get_attribute("src")
            if not img_url or "sack_work_end.png" in (img_url or ""):
                img_url = img_tag.get_attribute("data-src")
            item["이미지 URL"] = img_url if img_url else ""
        except Exception:
            item["이미지 URL"] = ""

        try:
            item["작가명"] = art.find_element(By.CLASS_NAME, "card-title").text.strip()
        except Exception:
            item["작가명"] = ""

        try:
            item["작품명"] = art.find_element(By.CLASS_NAME, "card-subtitle").text.strip()
        except Exception:
            item["작품명"] = ""

        try:
            desc_ps = art.find_elements(By.CSS_SELECTOR, "p.description span")
            if len(desc_ps) >= 2:
                item["재질"] = desc_ps[0].text.strip()
                item["사이즈 및 연도"] = desc_ps[1].text.strip()
            elif len(desc_ps) == 1:
                item["재질/사이즈"] = desc_ps[0].text.strip()
        except Exception:
            pass

        try:
            price_div = art.find_element(By.CLASS_NAME, "dotted")
            price_lis = price_div.find_elements(By.TAG_NAME, "li")
            for i, li in enumerate(price_lis):
                if "추정가" in li.text and i + 1 < len(price_lis):
                    item["추정가 (KRW)"] = price_lis[i + 1].text.strip()
                elif "시작가" in li.text and i + 1 < len(price_lis):
                    item["시작가 (KRW)"] = price_lis[i + 1].text.strip()
        except Exception:
            pass

        try:
            card_texts = art.find_elements(By.CLASS_NAME, "card-text")
            if len(card_texts) >= 2:
                item["마감 시간"] = card_texts[-1].text.strip()
        except Exception:
            pass

        try:
            link_tag = art.find_element(By.CLASS_NAME, "listimg")
            detail_url = link_tag.get_attribute("href")
            link = detail_url if detail_url else url
            item["바로가기 URL"] = f'=HYPERLINK("{link}", "{link}")'
        except Exception:
            item["바로가기 URL"] = f'=HYPERLINK("{url}", "{url}")'

        item["경매명"] = auction_title
        item["일정"] = auction_schedule
        item["전시장소"] = auction_location

        data_list.append(item)
        print(f"케이옥션 추출 완료: Lot {item.get('Lot', '')} - {item.get('작가명', '')}")

    return data_list


def scrape_kan_auction(driver):
    print("\n--- 칸옥션 공지 정보 수집 시작 ---")
    url = "http://www.kanauction.kr/auction/going/main"
    driver.get(url)

    time.sleep(3)

    data_list = []
    item = {}

    try:
        elements = driver.find_elements(By.TAG_NAME, "div")
        info_text = ""

        for el in elements:
            style = (el.get_attribute("style") or "").replace(" ", "")
            if style in ["text-align:center", "text-align:center;"]:
                txt = el.text.strip()
                if txt and ("경매" in txt or "칸옥션" in txt):
                    info_text = txt
                    break

        if not info_text:
            for el in elements:
                text = el.text.strip()
                if "칸옥션 제" in text and "미술품경매" in text and "예정" in text:
                    info_text = text
                    break

        item["공지 내용"] = info_text if info_text else "경매 정보를 찾을 수 없거나 아직 등록되지 않았습니다."
    except Exception as e:
        item["공지 내용"] = "에러 발생: " + str(e)

    item["바로가기 URL"] = f'=HYPERLINK("{url}", "{url}")'
    data_list.append(item)
    print("칸옥션 추출 완료.")

    return data_list


def scrape_myart_auction(driver):
    print("\n--- 마이아트옥션 공지 정보 수집 시작 ---")
    url = "https://myartauction.com/auctions/ongoing"
    driver.get(url)

    time.sleep(3)

    data_list = []
    item = {}

    try:
        page_source = driver.page_source
        if "NO CURRENT AUCTIONS" in page_source or "새로운 경매가 곧 시작됩니다" in page_source:
            item["공지 내용"] = "NO CURRENT AUCTIONS / 새로운 경매가 곧 시작됩니다"
        else:
            item["공지 내용"] = "현재 진행 중인 경매가 있습니다. (상세 내역은 홈페이지를 참고하세요)"
    except Exception as e:
        item["공지 내용"] = "에러 발생: " + str(e)

    item["바로가기 URL"] = f'=HYPERLINK("{url}", "{url}")'
    data_list.append(item)
    print("마이아트옥션 추출 완료.")

    return data_list


# =========================
# 3) eBay (Playwright 비동기 수집)
# =========================
async def async_scrape_ebay(keyword):
    print(f"\n--- 이베이(eBay) 정보 수집 시작 (검색어: {keyword}) ---")
    nkw = urllib.parse.quote_plus(keyword)
    url = f"https://www.ebay.com/sch/i.html?_nkw={nkw}&_sacat=0&_from=R40&_trksid=m570.l1313&_odkw={nkw}&_osacat=0"

    data_list = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            await page.goto(url, wait_until="networkidle", timeout=30000)
            for _ in range(3):
                await page.mouse.wheel(0, 1000)
                await asyncio.sleep(1)

            items = await page.query_selector_all("li.s-item, .s-card")
            for item in items:
                row = {}
                try:
                    title_el = await item.query_selector(".s-item__title, .s-card__title, div[role='heading']")
                    if not title_el:
                        continue
                    title = await title_el.inner_text()
                    row["항목명"] = title.replace("새 창 또는 새 탭에서 열림", "").strip()
                    if not row["항목명"] or row["항목명"] in ["Shop on eBay", "eBay 상품 더보기", "관련 상품"]:
                        continue
                except Exception:
                    continue

                try:
                    price_el = await item.query_selector(".s-item__price, .s-card__price, .s-item__price span")
                    row["가격 정보"] = await price_el.inner_text() if price_el else ""
                except Exception:
                    row["가격 정보"] = ""

                try:
                    shipping_el = await item.query_selector(
                        ".s-item__shipping, .s-item__logisticsCost, .s-card__shipping, .s-item__free-shipping"
                    )
                    if not shipping_el:
                        all_text = await item.inner_text()
                        if "배송" in all_text or "Shipping" in all_text:
                            shipping_el = await item.query_selector("span:has-text('배송'), span:has-text('Shipping')")
                    row["배송 정보"] = await shipping_el.inner_text() if shipping_el else "Shipping info not found"
                except Exception:
                    row["배송 정보"] = "Shipping info not found"

                try:
                    link_el = await item.query_selector(".s-item__link, .s-card__link, a")
                    href = await link_el.get_attribute("href")
                    row["바로가기 URL"] = f'=HYPERLINK("{href}", "{href}")' if href else ""
                except Exception:
                    row["바로가기 URL"] = ""

                try:
                    img_el = await item.query_selector(".s-item__image-img, .s-card__image-img, .s-card__link img, img")
                    img_url = await img_el.get_attribute("src") if img_el else ""
                    if not img_url or "placeholder" in img_url or "static" in img_url:
                        img_url = await img_el.get_attribute("data-src") if img_el else img_url
                    row["이미지 URL"] = img_url if img_url else ""
                except Exception:
                    row["이미지 URL"] = ""

                data_list.append(row)
                print(f"이베이 추출 완료: {row['항목명'][:30]}...")
        except Exception as e:
            print(f"이베이 수집 중 에러 발생: {e}")
        finally:
            await browser.close()

    return data_list


def scrape_ebay_auction(keyword):
    # Windows 로컬에서 실행할 때만 policy 적용 (Cloud는 linux라 영향 없음)
    if sys.platform == "win32":
        asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())
    return asyncio.run(async_scrape_ebay(keyword))


# =========================
# 4) Excel 저장 (IMAGE 함수로 URL 렌더링)
# =========================
def save_to_excel_with_images(data_dict, filename):
    wb = Workbook()
    first_sheet = True

    for sheet_name, data in data_dict.items():
        if first_sheet:
            ws = wb.active
            ws.title = sheet_name
            first_sheet = False
        else:
            ws = wb.create_sheet(title=sheet_name)

        df = pd.DataFrame(data)
        if df.empty:
            ws.append(["데이터가 없습니다."])
            continue

        # 헤더
        ws.append(list(df.columns))
        for c in range(1, len(df.columns) + 1):
            ws.cell(row=1, column=c).font = Font(bold=True)

        # 데이터
        for r_idx, row in enumerate(df.values, start=2):
            for c_idx, val in enumerate(row, start=1):
                col_name = df.columns[c_idx - 1]

                # 이미지 URL -> Excel IMAGE 함수로 표시
                if col_name in ["이미지 URL", "이미지"]:
                    img_url = val
                    if img_url and isinstance(img_url, str) and img_url.startswith("http"):
                        ws.cell(row=r_idx, column=c_idx).value = f'=_xlfn.IMAGE("{img_url}")'
                        ws.row_dimensions[r_idx].height = 80
                        ws.column_dimensions[get_column_letter(c_idx)].width = 25
                    else:
                        ws.cell(row=r_idx, column=c_idx).value = str(img_url) if img_url else ""
                else:
                    ws.cell(row=r_idx, column=c_idx).value = str(val) if val else ""

    wb.save(filename)
    print(f"파일 저장 완료: {filename}")


# =========================
# 5) main
# =========================
def main():
    print("브라우저를 초기화합니다...")

    try:
        driver = get_driver_cached()
    except Exception as e:
        print(f"브라우저 초기화 실패: {e}")
        raise e

    # 1) 사이트별 데이터 수집
    seoul_data = scrape_seoul_auction(driver)
    k_auction_data = scrape_k_auction(driver)
    kan_data = scrape_kan_auction(driver)
    myart_data = scrape_myart_auction(driver)

    # driver 종료
    try:
        driver.quit()
    except Exception:
        pass

    # 2) 이베이 데이터 수집
    ebay_keyword = "빈티지, 의약, 약국, 약, 약제상"
    ebay_data = scrape_ebay_auction(ebay_keyword)

    # 3) 엑셀 저장
    if seoul_data or k_auction_data or kan_data or myart_data or ebay_data:
        excel_filename = "auction_upcoming_hyperlinks.xlsx"
        data_dict = {
            "서울옥션": seoul_data,
            "케이옥션": k_auction_data,
            "칸옥션": kan_data,
            "마이아트옥션": myart_data,
            "이베이(eBay)": ebay_data,
        }

        save_to_excel_with_images(data_dict, excel_filename)
        print(f"\n>> 수집된 모든 데이터가 성공적으로 '{excel_filename}'에 탭으로 구분되어 저장되었습니다.")
    else:
        print("\n>> 수집된 데이터가 없어 엑셀 파일이 생성되지 않았습니다.")


if __name__ == "__main__":
    main()
