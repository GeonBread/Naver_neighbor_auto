from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

import pyperclip
import time
import schedule

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from datetime import datetime

import tkinter as tk
from tkinter import ttk, messagebox
import threading

# 서비스 계정 JSON 파일 경로
SERVICE_ACCOUNT_FILE = "aa.json"
SPREADSHEET_ID = "aa"
DEFAULT_SHEET_NAME = "Sheet1"

NAVER_ID = "aa"
NAVER_PASSWORD = "aa"

# Tkinter GUI 초기화 함수
def initialize_gui():
    global window, keyword_entry, message_entry, progress_bar, loading_label, progress_label, loading_window, collected_data
    width = 400
    height = 400

    window = tk.Tk()
    window.title("밤비봉 이웃 관리 프로그램")
    window.geometry(f"{width}x{height}")
    
    # 창을 윈도우 정중앙에 배치
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f"{width}x{height}+{x}+{y}")


    # 키워드 입력 섹션
    tk.Label(window, text="키워드 입력", font=("Arial", 12)).pack(pady=5)
    keyword_entry = tk.Entry(window, width=30, font=("Arial", 12))
    keyword_entry.pack(pady=5)

    tk.Button(window, text="이웃 정보 모으기", font=("Arial", 12), command=collect_blog_data).pack(pady=10)

    # 서로이웃 메시지 입력 섹션
    tk.Label(window, text="서로이웃 메시지 입력", font=("Arial", 12)).pack(pady=5)
    message_entry = tk.Text(window, height=5, width=30, font=("Arial", 12))
    message_entry.pack(pady=5)

    tk.Button(window, text="서로이웃 추가 시작", font=("Arial", 12), command=add_neighborhood).pack(pady=10)


# 로딩 창 초기화 함수
def initialize_loading_window():
    global loading_window, progress_bar, progress_label

    loading_window = tk.Toplevel(window)
    loading_window.title("로딩 중...")
    loading_window.geometry("300x200")
    loading_window.resizable(False, False)

    # 로딩 창을 윈도우 정중앙에 배치
    x = (window.winfo_screenwidth() // 2) - (300 // 2)
    y = (window.winfo_screenheight() // 2) - (200 // 2)
    loading_window.geometry(f"300x200+{x}+{y}")

    # 메인 윈도우 조작 금지
    loading_window.grab_set()

    # 로딩 창 닫기 방지
    loading_window.protocol("WM_DELETE_WINDOW", lambda: None)

    # 타이틀바는 유지하되 조작 버튼 제거
    loading_window.attributes('-toolwindow', True)
    
    tk.Label(loading_window, text="작업 진행 상황", font=("Arial", 12)).pack(pady=10)

    progress_bar = ttk.Progressbar(loading_window, orient="horizontal", length=250, mode="determinate")
    progress_bar.pack(pady=5)

    progress_label = tk.Label(loading_window, text="0% 완료", font=("Arial", 12))
    progress_label.pack(pady=5)


    # 중단 버튼
    tk.Button(loading_window, text="중단", command=stop_task).pack(pady=20)

# 블로그 데이터를 수집하는 함수
def collect_blog_data():
    
    global stop_flag
    stop_flag = False  # 중단 플래그 초기화

    #키워드 가져오기
    keyword = keyword_entry.get()
    if not keyword:
        messagebox.showerror("오류", "키워드를 입력하세요!")
        return

    initialize_loading_window()

    def scrape():
        try : 
            percent = 10
            blog_data = scrape_blog_data(driver, keyword)
            
            if blog_data is False:
                loading_window.destroy()
                messagebox.showinfo("완료", f"'{keyword}'에 대한 블로그가 없습니다!")
                return
            if not blog_data:
                loading_window.destroy()
                messagebox.showinfo("완료", f"작업이 중단되었습니다.")
                return

            progress_bar["value"] = percent
            progress_label.config(text=f"{percent}% 완료")

            blog_data = collect_additional_data(driver, blog_data)  
                
            #스프레드시트에 데이터 저장하고 마무리
            idx = get_next_empty_row_index(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME)
            write_to_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, blog_data, f"A{idx}")

            if stop_flag:
                loading_window.destroy()
                messagebox.showinfo("완료", f"작업된 부분까지만 저장했습니다.")
                return

            progress_bar["value"] = 100
            progress_label.config(text=f"{percent}% 완료")
            
            loading_window.destroy()
            messagebox.showinfo("완료", f"'{keyword}'에 대한 이웃 정보가 수집되었습니다!")
        except Exception as e:
            loading_window.destroy()
            print(e)
            messagebox.showerror("오류", f"작업 중 문제가 발생했습니다: {str(e)}")

    threading.Thread(target=scrape).start()

# 서로 이웃 추가 시작 함수
def add_neighborhood():
    
    global stop_flag
    stop_flag = False  # 중단 플래그 초기화

    message = message_entry.get("1.0", tk.END).strip()
    if not message:
        messagebox.showerror("오류", "메시지를 입력하세요!")
        return

    messagebox.showinfo("완료", f"다음 메시지로 서로이웃 신청을 시작합니다:\n\n{message}")

    initialize_loading_window()

    def scrape():
        links = filter_and_transform_links(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME)

        # 서로이웃 메시지
        message = message_entry.get("1.0", "end-1c")

        # 서로이웃 추가 실행
        start_index = 0

        while start_index != -1:
            print(f"Starting from index: {start_index}")
            start_index = add_neighbors(driver, links[start_index:], message, max_count=100)

            if start_index >= -3 and start_index <= -1 :
                break
        
        loading_window.destroy()
        if start_index == -1:
            messagebox.showinfo("완료", f"이웃 추가 완료했습니다")
        if start_index == -2: #하루 신청 수 100명 채운 경우
            messagebox.showinfo("완료", f"오늘 하루 신청 수 다 채웠습니다!");
        if start_index == -3: #중단한 경우
            messagebox.showinfo("완료", f"작업 중단하였습니다.")

    threading.Thread(target=scrape).start()

# 확정 로딩바 작업 함수
def start_loading_task():
    initialize_loading_window()

    def background_task():
        for i in range(101):
            time.sleep(0.05)  # 작업 시뮬레이션
            progress_bar["value"] = i
            progress_label.config(text=f"{i}% 완료")
        loading_window.destroy()
        messagebox.showinfo("완료", "작업이 완료되었습니다!")

    threading.Thread(target=background_task).start()


def stop_task():
    global stop_flag
    if messagebox.askyesno("중단 확인", "작업을 중단하시겠습니까?"):
        stop_flag = True  # 중단 플래그 설정


#네이버 로그인하는 함수
def naver_login(driver, naver_id, naver_pw):
    """
    Selenium을 사용해 네이버에 로그인하는 함수입니다.

    매개변수:
    - driver: Selenium WebDriver 인스턴스.
    - naver_id: 사용자의 네이버 ID.
    - naver_pw: 사용자의 네이버 비밀번호.
    """
    url = 'https://nid.naver.com/nidlogin.login?mode=form&url=https://blog.naver.com/'
    driver.get(url)

    # ID 입력 필드 대기 후 복사하여 붙여넣기
    id_field = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "id"))
    )
    id_field.click()
    clipboard_backup = pyperclip.paste()  # 기존 클립보드 내용을 백업
    pyperclip.copy(naver_id)  # ID를 클립보드로 복사
    
    actions = webdriver.ActionChains(driver)
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    time.sleep(1)  # 봇 탐지를 피하기 위해 대기

    # 비밀번호 입력 필드 대기 후 복사하여 붙여넣기
    pw_field = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "pw"))
    )
    pw_field.click()
    pyperclip.copy(naver_pw)  # 비밀번호를 클립보드로 복사
    actions.key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    time.sleep(1)  # 봇 탐지를 피하기 위해 대기

    # 기존 클립보드 내용을 복원
    pyperclip.copy(clipboard_backup)

    # 로그인 버튼 클릭
    login_button = WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.ID, "log.login"))
    )
    login_button.click()

    print("로그인 프로세스가 시작되었습니다.")
    
#기다렸다가 요소가 보이면 바로 클릭하는 함수
def wait_and_click(driver, xpath, timeout=1):
    """
    WebDriverWait를 사용하여 지정된 XPATH의 요소가 클릭 가능할 때까지 기다린 후 클릭합니다.

    Args:
        driver (WebDriver): Selenium WebDriver 인스턴스.
        xpath (str): 클릭할 요소의 XPATH.
        timeout (int): 요소를 기다릴 최대 시간(초).

    Returns:
        bool: 클릭 성공 여부.

    Raises:
        None: 실패 시 로그만 출력하고 False를 반환.
    """
    try:
        # 요소가 클릭 가능해질 때까지 대기 후 클릭
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        element.click()
        print(f"성공적으로 클릭: {xpath}")
        return True
    except TimeoutException:
        print(f"Timeout: {timeout}초 내에 요소를 클릭할 수 없습니다. XPATH: {xpath}")
    except NoSuchElementException:
        print(f"요소를 찾을 수 없습니다. XPATH: {xpath}")
    except Exception as e:
        print(f"예상치 못한 오류 발생: {e}")
    return False

# 구글 스프레드시트 초기화 함수
def initialize_service():
    credentials = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE,
        scopes=["https://www.googleapis.com/auth/spreadsheets"]
    )
    return build("sheets", "v4", credentials=credentials)

def initialize_sheet(service, spreadsheet_id, sheet_name):
    # 헤더 추가
    header = [
        ["닉네임", "포스팅 제목", "링크", "오늘 방문자수", "전체 방문자수", "이웃수", "최근 글 발행 시간", 
         "서로이웃 신청 여부", "", "날짜", "서로이웃 신청 수"]
    ]
    range_name = f"{sheet_name}!A1:K1"  # A1부터 K1까지의 범위 지정
    body = {
        "values": header
    }

    # 헤더 업데이트
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="RAW",
        body=body
    ).execute()

    # 서식 설정 요청 생성
    requests = [
        # 필터링 기능 추가 요청
        {
            "setBasicFilter": {
                "filter": {
                    "range": {
                        "sheetId": 0,  # 기본 시트 ID
                        "startRowIndex": 0,  # 헤더 포함
                        "startColumnIndex": 3,  # D 열 시작
                        "endColumnIndex": 8   # H 열 끝 (H + 1)
                    }
                }
            }
        },
        # D, E, F 열: 숫자에 천 단위 콤마 추가
        {
            "repeatCell": {
                "range": {
                    "sheetId": 0,  # 기본 시트 ID
                    "startRowIndex": 1,  # 헤더 이후부터 적용
                    "startColumnIndex": 3,  # D 열
                    "endColumnIndex": 6   # F 열
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": "#,##0"
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        },
        # F 열: "명" 추가
        {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": 1,
                    "startColumnIndex": 5,  # F 열
                    "endColumnIndex": 6
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "NUMBER",
                            "pattern": "#,##0명"
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        },
        # G 열: 날짜 서식 설정
        {
            "repeatCell": {
                "range": {
                    "sheetId": 0,
                    "startRowIndex": 1,
                    "startColumnIndex": 6,  # G 열
                    "endColumnIndex": 7
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": "DATE",
                            "pattern": "yyyy. MM. dd"
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        }
    ]

    # 서식 요청 실행
    service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={"requests": requests}
    ).execute()


# 마지막 데이터 행의 다음 행 인덱스를 반환하는 함수
def get_next_empty_row_index(service, spreadsheet_id, sheet_name):
    # A열부터 데이터를 읽어와야 데이터가 있는 마지막 행을 알 수 있음
    data = read_from_sheet(service, spreadsheet_id, sheet_name, "A:A")  # A열 데이터만 읽기
    if not data:  # 데이터가 없으면 첫 번째 행 이후부터 시작
        return 2
    return len(data) + 1  # 마지막 데이터 행의 다음 행 인덱스 반환

#서로이웃 신청 횟수 카운팅 하기 위한 함수랄까..
def initialize_today_neighbors(service, spreadsheet_id, sheet_name):
    """
    오늘 날짜에 해당하는 서로이웃 신청 수를 초기화하거나 확인하는 함수.
    
    매개변수:
    - service: Google Sheets API 서비스 객체.
    - spreadsheet_id: 스프레드시트 ID.
    - sheet_name: 처리할 시트 이름.
    
    반환값:
    - 행의 위치 (row_index, 1부터 시작)
    - 오늘 날짜에 해당하는 서로이웃 신청 수 값 (K열 값)
    """
    today_date = datetime.now().strftime("%Y-%m-%d")
    data = read_from_sheet(service, spreadsheet_id, sheet_name, "J:K")
    
    for i, row in enumerate(data[1:], start=2):  # 헤더를 건너뛰고 2번째 행부터 시작
        j_value = row[0] if len(row) > 0 else ""  # J열 값 확인
        k_value = int(row[1]) if len(row) > 1 and row[1].isdigit() else 0  # K열 값 확인
        
        if j_value == today_date:  # 오늘 날짜와 동일한 값 발견
            return i, k_value
    
    # 오늘 날짜가 없는 경우 새 행의 위치와 0 반환
    idx = len(data) + 1
    write_to_sheet(service, spreadsheet_id, sheet_name, [[today_date]], f"J{idx}")
    
    return idx, 0

# 스프레드시트에 데이터 쓰기
def write_to_sheet(service, spreadsheet_id, sheet_name, data, start_cell="A2"):
    """
    Google Sheets에 데이터를 작성하는 함수

    Args:
        service: Google Sheets API 서비스 객체
        spreadsheet_id: 스프레드시트 ID
        sheet_name: 시트 이름
        data: 입력할 데이터 (2D 리스트)
        start_cell: 데이터 쓰기를 시작할 셀 위치 (기본값: A2)
    """
    range_name = f"{sheet_name}!{start_cell}"
    body = {
        "values": data
    }

    # USER_ENTERED를 사용하여 데이터 자동 형식 감지
    service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id,
        range=range_name,
        valueInputOption="USER_ENTERED",  # 데이터 자동 형식 감지
        body=body
    ).execute()

# 스프레드시트에서 데이터 읽기
def read_from_sheet(service, spreadsheet_id, sheet_name, range_columns):
    range_name = f"{sheet_name}!{range_columns}"
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    rows = result.get("values", [])
    return rows

# C열과 H열만 추출하는 함수
def extract_columns(data, column_indices):
    extracted = []
    for row in data:
        extracted.append([row[i] if i < len(row) else "" for i in column_indices])
    return extracted

# Selenium 드라이버 설정
def setup_driver():
    options = webdriver.ChromeOptions()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    return webdriver.Chrome(options=options)

# 블로그 검색 및 데이터 수집
def scrape_blog_data(driver, keyword):
    blog_data = []
    base_url = "https://search.naver.com/search.naver?ssc=tab.blog.all&sm=tab_jum&query="
    search_url = base_url + keyword
    driver.get(search_url)
    time.sleep(2)

    # 페이지 스크롤
    for _ in range(10):
        #중단 요청 시 루프 종료
        if stop_flag:
            return []
        
        driver.execute_script("window.scrollBy(0, 3000);")
        time.sleep(0.5)

    results = driver.find_elements(By.XPATH, '//*[@id="main_pack"]/section/div[1]/ul/li[@class="bx"]')

    if not results:
        return False

    for result in results:
        #중단 요청 시 루프 종료
        if stop_flag:
            return []
        try:
            nickname = result.find_element(By.XPATH, './/div/div[1]/div[2]/div/a').text
            title_element = result.find_element(By.XPATH, './/div[@class="title_area"]/a')
            title_text = title_element.text
            title_href = title_element.get_attribute('href')

            if "blog.naver.com" in title_href:
                blog_data.append([nickname, title_text, title_href])
        except Exception:
            continue

    return blog_data

# 블로그 추가 데이터 수집
def collect_additional_data(driver, blog_data):
    percent = 10
    addPercent = 85 / len(blog_data)
    today = datetime.now().strftime("%Y. %m. %d")  # 오늘 날짜 서식
    
    for index, entry in enumerate(blog_data):
        #중단 요청 시 루프 종료
        if stop_flag:
            return blog_data[0:index]

        percent += addPercent
        progress_bar["value"] = percent
        progress_label.config(text=f"{percent:.2f}% 완료")

        link = entry[2]
        try:
            parts = link.split('/')
            mobile_main_blog = f"https://m.blog.naver.com/{parts[3]}"
            driver.get(mobile_main_blog)

            try:
                today_visitors_text = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@class="cover_content__OApzT"]//div[@class="count__T3YO8"]'))
                ).text
                today_visitors, all_visitors = today_visitors_text.split('전체')
                today_visitors = int(today_visitors.replace("오늘", "").replace(",", "").strip())
                all_visitors = int(all_visitors.replace(",", "").strip())
            except Exception:
                today_visitors, all_visitors = 0, 0


            # 이웃 수 데이터 가져오기
            try:
                neighbors_text = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@class="bloger_area__cmYsI"]//span[@class="buddy__fw6Uo"]'))
                ).text
                neighbors = int(neighbors_text.replace("명의 이웃", "").replace(",", "").strip())
            except Exception:
                neighbors = 0

            # 최근 글 발행 시간 가져오기
            try:
                last_post_time = WebDriverWait(driver, 1).until(
                    EC.presence_of_element_located((By.XPATH, '//span[@class="time__mHZOn"]'))
                ).text

                if "시간 전" in last_post_time or "분 전" in last_post_time:
                    last_post_time = today  # "xx시간 전" 또는 "xx분 전"이면 오늘 날짜로 설정
                elif "." in last_post_time:  # "년도. 월. 일." 형식인지 확인
                    try:
                        # 날짜 형식 검증 및 변환
                        parsed_date = datetime.strptime(last_post_time.strip(), "%Y. %m. %d.")
                        last_post_time = parsed_date.strftime("%Y. %m. %d")  # 표준화된 형식으로 저장
                    except ValueError:
                        last_post_time = today  # 예상치 못한 형식이면 기본값으로 오늘 날짜 설정
            except Exception:
                last_post_time = "-"

            entry.extend([today_visitors, all_visitors, neighbors, last_post_time])
        except Exception:
            continue

    return blog_data


def filter_and_transform_links(service, spreadsheet_id, sheet_name):
    """
    스프레드시트에서 C열(링크)과 H열(신청 여부)을 읽어 필터링 후 반환합니다.
    """
    data = read_from_sheet(service, spreadsheet_id, sheet_name, "A:H")  # A~H 열 읽기
    filtered_links = []
    for idx, row in enumerate(data[1:], start = 2):  # 첫 행(헤더) 제외
        if len(row) < 8:  # H열(신청 여부)에 데이터가 없으면 (동그라미든 엑스든 상관없이)
            mobile_link = convert_to_mobile_link(row[2])  # C열(링크) 변환
            filtered_links.append([row[0], mobile_link, idx])  # 닉네임(A열), 모바일 링크 반환
    return filtered_links


def convert_to_mobile_link(link):
    """
    PC 링크를 모바일 링크로 변환합니다.
    """
    parts = link.split('/')
    return f"https://m.blog.naver.com/{parts[3]}"


def add_neighbors(driver, links, message, max_count=100):
    """
    Selenium을 사용해 서로이웃 추가를 수행하며, 자정을 감지하여 작업을 중단한 뒤 이어서 실행합니다.
    
    매개변수:
    - driver: Selenium WebDriver 인스턴스.
    - links: [(닉네임, 링크)] 형식의 리스트.
    - message: 서로이웃 신청 메시지.
    - max_count: 하루 최대 작업 개수.
    """
    global stop_flag

    row_index, count = initialize_today_neighbors(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME)
    percent_unit = 100 / (len(links) + 1)

    for i, (nickname, link, idx) in enumerate(links):        
        time.sleep(1)
        #중단 요청 시 루프 종료
        if stop_flag:
            return -3
        
        percent = percent_unit * (i + 1)
        progress_bar["value"] = percent
        progress_label.config(text=f"{percent:.2f}% 완료")

        if count >= max_count:
            print("Reached maximum count for today.")
            return -2

        # 자정 감지
        now = datetime.now()
        if now.hour == 0 and now.minute == 0:  # 자정 확인
            print("Midnight detected. Stopping process temporarily.")
            return i  # 중단된 작업의 인덱스를 반환

        driver.get(link)

        try:
            # 이웃추가 버튼 클릭
            wait_and_click(driver, '//*[@id="root"]/div[4]/div/div[3]/div[1]/button')
            # 서로이웃 옵션 클릭
            if wait_and_click(driver, '//*[@id="bothBuddyRadio"]'):
                # 메시지 입력
                text_area = driver.find_element(By.XPATH, '//*[@id="buddyAddForm"]/fieldset/div/div[2]/div[3]/div/textarea')
                text_area.send_keys(message)
                # 확인 버튼 클릭
                wait_and_click(driver, '/html/body/ui-view/div[2]/a[2]')
                count += 1
            
            
            # 이미 신청했거나 막혀있는 경우
            else :
                try: 
                    #서로이웃 더 이상 추가가 안 된다면 그날 카운팅을 100으로 초기화하고 함수 종료.
                    _popup_message = driver.find_element(By.XPATH, '//*[@id="root"]/div[7]/div/div/div/div[1]/p').text
                    if (_popup_message !=  '서로이웃 신청 진행중입니다. 서로이웃\n신청을 취소하시겠습니까?' and
                        _popup_message != '상대방이 이웃 관계를\n제한하여, 이웃/서로이웃을\n신청하실 수 없습니다.'):
                        write_to_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, [[100]], f"K{row_index}")
                        return -1
                except Exception as e:
                    try:
                        if driver.find_element(By.XPATH, '//*[@id="ct"]/fieldset/div/div[2]/div/span[1]/label').text \
                            == '이웃을 서로이웃으로 변경합니다.':
                            update_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, nickname, "이웃", count, row_index, idx)
                        elif driver.find_element(By.XPATH, '//*[@id="ct"]/fieldset/div/div[2]/div/span[1]/label').text \
                            == '서로이웃을 이웃으로 변경합니다.':
                            update_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, nickname, "서로 이웃", count, row_index, idx)
                    except Exception as e:
                        update_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, nickname, "신청 불가", count, row_index, idx)
                        print(e)
                    continue

            # 스프레드시트 갱신
            update_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME, nickname, "O", count, row_index, idx)
        except Exception as e:
            print(f"Error with {nickname}: {e}")
            continue

    return -1  # 작업 완료 시 -1 반환


def update_sheet(service, spreadsheet_id, sheet_name, nickname, status, count, row_index, idx):
    """
    신청 여부(H열), 날짜(J열), 서로이웃 신청 수(K열)를 갱신합니다.

    매개변수:
    - service: Google Sheets API 서비스 객체.
    - spreadsheet_id: 스프레드시트 ID.
    - sheet_name: 업데이트할 시트 이름.
    - nickname: 업데이트할 닉네임.
    - status: H열에 기록할 상태 값 (예: "O").
    - count: 오늘의 서로이웃 신청 카운팅 (K열에 기록).
    """

    write_to_sheet(service, spreadsheet_id, sheet_name, [[status]], f"H{idx}")
    print(f"Updated {nickname} status to {status} in H{idx}.")

    # J열 확인 및 갱신   
    write_to_sheet(service, spreadsheet_id, sheet_name, [[count]], f"K{row_index}")

# 스프레드시트 저장 및 읽기
if __name__ == "__main__":

    #웹드라이버 설정
    driver = setup_driver()

    #구글 스프레드 기본 초기화 하기
    service = initialize_service()
    initialize_sheet(service, SPREADSHEET_ID, DEFAULT_SHEET_NAME)

    # 네이버 로그인 함수 호출
    naver_login(driver, NAVER_ID, NAVER_PASSWORD)

    initialize_gui()
    window.mainloop()