import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.styles import Border, Side
import os
import time
from selenium.webdriver.common.keys import Keys

# 설정 변수
CONFIG = {
    'EXCEL_FILE_NAME': '제주도 기관 공고.xlsx',  # 엑셀 파일명
    'SHEET_NAME': '일자리정보',                 # 시트명
    'MAX_PAGES': 30,                           # 최대 페이지 수
    'START_ROW': 5,                           # 데이터 시작 행
    'ROW_HEIGHT': 88,                         # 행 높이
    'URL': 'https://jeju.work.go.kr/main.do/empInfo/empInfoSrch/list/retriveWorkRegionEmpCodeList.do'  # 대상 URL
}

def setup_driver():
    """Selenium 웹드라이버 설정"""
    try:
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')  # 헤드리스 모드 활성화
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-extensions')
        options.add_argument('--disable-infobars')
        options.add_argument('--window-size=1920,1080')  # 화면 크기 설정
        
        # 웹드라이버 자동 설치 및 설정
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=options)
        return driver
    except Exception as e:
        print(f"웹드라이버 설정 중 오류 발생: {str(e)}")
        raise

def extract_table_data(url):
    """웹페이지에서 테이블 데이터 추출"""
    driver = setup_driver()
    all_data = []
    headers = ["일련번호", "회사/공고/상세", "근무조건", "마감/등록일"]
    
    try:
        # 접속 시도 (1초당 5회, 최대 5회)
        max_retries = 5
        retry_count = 0
        connected = False
        
        while retry_count < max_retries and not connected:
            try:
                print(f"웹페이지 접속 시도 중... (시도 {retry_count + 1}/{max_retries})")
                driver.get(url)
                wait = WebDriverWait(driver, 20)
                
                # iframe 확인
                try:
                    iframe = wait.until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))
                    driver.switch_to.frame(iframe)
                    time.sleep(3)
                    connected = True
                    print("웹페이지 접속 성공")
                except:
                    print("iframe을 찾을 수 없습니다.")
                    retry_count += 1
                    time.sleep(1)
                    continue
                    
            except Exception as e:
                print(f"접속 실패: {str(e)}")
                retry_count += 1
                time.sleep(1)
                continue
        
        if not connected:
            print("최대 재시도 횟수 초과. 프로그램을 종료합니다.")
            return pd.DataFrame()

        sequence_number = 1
        for page in range(1, CONFIG['MAX_PAGES'] + 1):
            print(f"{page}페이지 처리 중...")
            if page > 1:
                try:
                    # 페이지 이동 전 테이블이 로드될 때까지 대기
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.board-list")))
                    time.sleep(2)
                    
                    # JavaScript로 페이지 이동
                    driver.execute_script(f"fn_Search({page});")
                    time.sleep(3)
                    
                    # 페이지 이동 후 테이블이 다시 로드될 때까지 대기
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.board-list")))
                    time.sleep(2)
                except Exception as e:
                    print(f"{page}페이지 이동 중 오류 발생: {str(e)}")
                    continue

            try:
                # 테이블 찾기
                table = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table.board-list")))
                time.sleep(2)
                
                # 테이블 행 가져오기
                rows = table.find_elements(By.TAG_NAME, "tr")
                if not rows or len(rows) <= 1:
                    print(f"{page}페이지에 데이터가 없습니다.")
                    continue

                for row in rows[1:]:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if not cells or len(cells) < 5:
                        continue
                    
                    # 1. 일련번호
                    row_data = [str(sequence_number)]
                    sequence_number += 1
                    
                    # 2. 회사/공고/상세
                    try:
                        company_name = cells[1].find_element(By.CSS_SELECTOR, "a.cp_name").text.strip()
                    except:
                        company_name = cells[1].text.strip()
                    
                    try:
                        job_title = cells[2].find_element(By.CSS_SELECTOR, "div.cp-info-in a").text.strip()
                    except:
                        job_title = cells[2].text.strip()
                    
                    details = []
                    try:
                        job_description = cells[2].find_element(By.CSS_SELECTOR, "p.mt10").text.strip()
                        if job_description:  # 빈 문자열이 아닌 경우만 추가
                            details.append(job_description)
                    except:
                        pass
                    
                    try:
                        info_p = cells[2].find_element(By.CSS_SELECTOR, "p:not(.mt10)")
                        em_elements = info_p.find_elements(By.TAG_NAME, "em")
                        for em in em_elements:
                            text = em.text.strip()
                            if text:  # 빈 문자열이 아닌 경우만 추가
                                details.append(text)
                    except:
                        pass
                    
                    # 회사/공고/상세 합치기
                    col2 = f"{company_name}\n{job_title}"
                    if details:
                        col2 += "\n" + "\n".join(details)
                    row_data.append(col2)
                    
                    # 3. 근무조건
                    try:
                        work_conditions = cells[3].find_element(By.CSS_SELECTOR, "div.cp-info").text.strip()
                    except:
                        work_conditions = cells[3].text.strip()
                    row_data.append(work_conditions)
                    
                    # 4. 마감/등록일
                    try:
                        deadline = cells[4].find_element(By.CSS_SELECTOR, "div.cp-info").text.strip()
                    except:
                        deadline = cells[4].text.strip()
                    row_data.append(deadline)
                    
                    # 데이터 유효성 검사
                    if all(row_data):  # 모든 필드가 비어있지 않은 경우만 추가
                        all_data.append(row_data)
                
                print(f"{page}페이지 데이터 추출 완료")
            except Exception as e:
                print(f"{page}페이지 데이터 추출 중 오류 발생: {str(e)}")
                continue

        if not all_data:
            print("데이터를 추출할 수 없습니다.")
            return pd.DataFrame()

        try:
            driver.switch_to.default_content()
        except:
            pass

        df = pd.DataFrame(all_data, columns=headers)
        return df

    except Exception as e:
        print(f"데이터 추출 중 오류 발생: {str(e)}")
        return pd.DataFrame()
    finally:
        driver.quit()

def save_to_excel(df, output_path):
    """데이터프레임을 엑셀 파일로 저장 (일자리정보 시트만 업데이트)"""
    try:
        if os.path.exists(CONFIG['EXCEL_FILE_NAME']):
            print("기존 엑셀 파일 로드 중...")
            wb = openpyxl.load_workbook(CONFIG['EXCEL_FILE_NAME'])
            
            # 일자리정보 시트가 있으면 해당 시트만 업데이트
            if CONFIG['SHEET_NAME'] in wb.sheetnames:
                ws = wb[CONFIG['SHEET_NAME']]
                # START_ROW부터 데이터 삭제 (이전 행은 유지)
                for row in range(CONFIG['START_ROW'], ws.max_row + 1):
                    for col in range(2, ws.max_column + 2):
                        ws.cell(row=row, column=col).value = None
            else:
                # 일자리정보 시트가 없으면 새로 생성
                ws = wb.create_sheet(CONFIG['SHEET_NAME'])
        else:
            print("새 엑셀 파일 생성 중...")
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = CONFIG['SHEET_NAME']

        # 테두리 스타일 정의
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # START_ROW부터 데이터 입력
        for i, row in df.iterrows():
            for j, value in enumerate(row):
                cell = ws.cell(row=CONFIG['START_ROW']+i, column=2+j, value=value)
                # 자동 줄바꿈 설정
                cell.alignment = openpyxl.styles.Alignment(wrap_text=True)
                # 테두리 적용
                cell.border = thin_border
            # 행 높이 설정
            ws.row_dimensions[CONFIG['START_ROW']+i].height = CONFIG['ROW_HEIGHT']

        print("엑셀 파일 저장 중...")
        wb.save(CONFIG['EXCEL_FILE_NAME'])
        print(f"데이터가 '{CONFIG['EXCEL_FILE_NAME']}' 파일의 '{CONFIG['SHEET_NAME']}' 시트에 저장되었습니다.")
        
    except Exception as e:
        print(f"엑셀 파일 저장 중 오류 발생: {str(e)}")
        raise

def main():
    """메인 실행 함수"""
    try:
        print("데이터 추출 중...")
        df = extract_table_data(CONFIG['URL'])
        
        if df.empty:
            print("데이터를 추출할 수 없습니다.")
            return

        print("엑셀 파일 저장 중...")
        save_to_excel(df, CONFIG['EXCEL_FILE_NAME'])
        print("데이터 업데이트가 완료되었습니다.")

    except Exception as e:
        print(f"오류 발생: {str(e)}")

if __name__ == "__main__":
    main() 