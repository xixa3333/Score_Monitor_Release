import time
import smtplib
import configparser
import base64
import os
import sys
import logging
import json
import traceback
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# =================日誌設定=================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("system_log.txt", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

# ================= 寄件者設定 =================
ENCRYPTED_SMTP_PASS = ""  # 請使用者自行在 config.txt 設定寄件者密碼
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "3333xixa3333@gmail.com"
SENDER_NAME = "高科成績通知系統"
HISTORY_FILE = "grade_history.json"

def get_credentials():
    config = configparser.ConfigParser()
    if not os.path.exists('config.txt'):
        logging.error("找不到 config.txt 設定檔！")
        sys.exit()
    try:
        config.read('config.txt', encoding='utf-8')
        return config['User']['Student_ID'], config['User']['Student_Password'], config['User']['Target_Email']
    except Exception as e:
        logging.error(f"讀取 config.txt 失敗: {e}")
        sys.exit()

def decode_password(encoded_str):
    try:
        return base64.b64decode(encoded_str).decode("utf-8")
    except:
        return ""

# ================= 成績紀錄功能 =================
def load_history():
    if not os.path.exists(HISTORY_FILE):
        return {}
    try:
        with open(HISTORY_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except:
        return {}

def save_history(grades):
    try:
        with open(HISTORY_FILE, 'w', encoding='utf-8') as f:
            json.dump(grades, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logging.error(f"無法儲存成績紀錄: {e}")

# ================= 郵件發送 (已修改評語邏輯) =================
def send_grade_update_email(target_email, new_grades):
    subject = "【成績通知】有新的成績公布了！"
    rows_html = ""
    
    for subject_name, score_text in new_grades:
        comment = ""
        score_color = "black"
        
        # 嘗試將分數轉為數字以進行判斷
        try:
            score_val = float(score_text)
            if score_val < 60:
                comment = "不好意思老師這次撈不動 😭"
                score_color = "red"
            else:
                comment = "恭喜你被老師撈撈上岸了 🎉"
                score_color = "green"
        except ValueError:
            # 如果分數不是數字 (例如: "通過", "抵免"), 就不顯示評語
            comment = "" 
            score_color = "blue"

        rows_html += f"""
        <tr>
            <td style='padding:8px;border:1px solid #ddd;'>{subject_name}</td>
            <td style='padding:8px;border:1px solid #ddd;color:{score_color};font-weight:bold;'>{score_text}</td>
            <td style='padding:8px;border:1px solid #ddd;'>{comment}</td>
        </tr>
        """

    content = f"""
    <h3>帥哥/美女你好：</h3>
    <p>系統偵測到下列科目已有分數：</p>
    <table style='border-collapse: collapse; width: 100%; max-width: 600px;'>
        <tr style='background-color: #f2f2f2;'>
            <th style='padding:8px;border:1px solid #ddd;text-align:left;'>科目名稱</th>
            <th style='padding:8px;border:1px solid #ddd;text-align:left;'>分數</th>
            <th style='padding:8px;border:1px solid #ddd;text-align:left;'>系統評語</th>
        </tr>
        {rows_html}
    </table>
    <br>
    <p><a href='https://stdsys.nkust.edu.tw/'>點此前往校務系統</a></p>
    """
    msg = MIMEText(content, 'html', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = formataddr((Header(SENDER_NAME, 'utf-8').encode(), SMTP_USER))
    msg['To'] = target_email

    try:
        real_password = decode_password(ENCRYPTED_SMTP_PASS)
        server = smtplib.SMTP(SMTP_HOST, SMTP_PORT)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(SMTP_USER, real_password)
        server.send_message(msg)
        server.quit()
        logging.info(f"✅ 已發送成績通知郵件 (共 {len(new_grades)} 科)")
    except Exception as e:
        logging.error(f"❌ 郵件發送失敗: {e}")

def parse_current_grades(driver):
    grades = {}
    try:
        rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
        for row in rows:
            cols = row.find_elements(By.TAG_NAME, "td")
            if len(cols) >= 6:
                subj_name = cols[2].text.strip()
                score_text = cols[5].text.strip()
                grades[subj_name] = score_text
    except StaleElementReferenceException:
        return {}
    except Exception:
        return {}
    return grades

# ================= 單次執行任務 (核心邏輯) =================
def run_browser_task(nkust_id, nkust_pwd, target_email):
    """
    這是一個會執行「開啟瀏覽器 -> 登入 -> 監控」的函式。
    如果需要重啟，這個函式會結束並回傳 'RESTART'。
    """
    # 初始化歷史紀錄
    history_grades = load_history()
    
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1200,900") # 設定一個舒服的視窗大小
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--log-level=3")

    driver = None
    try:
        logging.info("🚀 啟動瀏覽器視窗...")
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        
        logging.info("前往登入頁面...")
        driver.get("https://stdsys.nkust.edu.tw/student/Account/Login")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "usernameOrEmailAddress"))
        ).send_keys(nkust_id)
        driver.find_element(By.NAME, "Password").send_keys(nkust_pwd)

        time.sleep(3) 
        
        logging.info("🔄 進入自動登入迴圈 (持續嘗試驗證)...")
        start_time = time.time()
        timeout_seconds = 120 

        while True:
            # 1. 【檢查成功條件】：如果側邊選單 (bi-list) 出現，代表登入成功
            if len(driver.find_elements(By.CLASS_NAME, "bi-list")) > 0:
                logging.info("✅ 登入成功！(偵測到選單)")
                break
            
            # 2. 【檢查超時】：如果超過時間還沒進去，就重啟
            if time.time() - start_time > timeout_seconds:
                logging.error(f"❌ 登入超時 (超過 {timeout_seconds} 秒)，準備重啟...")
                return "RESTART"

            # 3. 【動作】：嘗試點擊「登入」按鈕
            try:
                login_btns = driver.find_elements(By.ID, "LoginButton")
                if len(login_btns) > 0 and login_btns[0].is_displayed():
                    login_btns[0].click()
                else:
                    sub_btns = driver.find_elements(By.CSS_SELECTOR, "button[type='submit']")
                    if len(sub_btns) > 0:
                        sub_btns[0].click()
            except Exception:
                pass 

            # 4. 【動作】：嘗試點擊「驗證 OK」彈窗
            try:
                ok_btns = driver.find_elements(By.CSS_SELECTOR, "button.swal2-confirm")
                if len(ok_btns) > 0 and ok_btns[0].is_displayed():
                    logging.info("✨ 偵測到驗證彈窗，點擊 OK...")
                    ok_btns[0].click()
                    time.sleep(3) 
            except Exception:
                pass

            time.sleep(1)

        # 進入成績頁面
        menu_btn = driver.find_element(By.CLASS_NAME, "bi-list")
        driver.execute_script("arguments[0].click();", menu_btn) 
        time.sleep(1)

        try:
            score_menu = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//p[contains(text(), '成績查詢')]"))
            )
            score_menu.click()
            time.sleep(1) 
        except:
            pass

        target_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, 'a[href="/student/Score/PresentSemester"]'))
        )
        target_link.click()

        logging.info(f"🟢 開始監控 (視窗模式)...")
        
        check_count = 0
        while True:
            try:
                # 檢查瀏覽器是否還活著
                _ = driver.title 
            except WebDriverException:
                logging.error("⚠️ 瀏覽器視窗似乎被關閉了。準備重啟...")
                return "RESTART"

            # 檢查是否被登出 (URL 包含 Login)
            if "Login" in driver.current_url:
                logging.warning("⚠️ 偵測到被系統自動登出！")
                logging.warning("🔄 正在準備重新啟動瀏覽器並登入...")
                return "RESTART" # 回傳重啟訊號，跳出函式

            # 正常的監控邏輯
            page_source = driver.page_source
            check_count += 1
            
            if "本學期尚無送達成績資料" in page_source:
                if check_count % 10 == 0: 
                    logging.info(f"[{time.strftime('%H:%M:%S')}] 尚無資料...")
            else:
                current_grades = parse_current_grades(driver)
                if not current_grades:
                    pass
                else:
                    new_updates = []
                    for subj, score in current_grades.items():
                        if score != "成績未送達":
                            old_score = history_grades.get(subj)
                            if old_score != score:
                                new_updates.append((subj, score))
                                history_grades[subj] = score
                    
                    if new_updates:
                        logging.info(f"🚨 發現 {len(new_updates)} 科新成績！")
                        send_grade_update_email(target_email, new_updates)
                        save_history(history_grades)
                    else:
                        if check_count % 10 == 0:
                            logging.info(f"[{time.strftime('%H:%M:%S')}] 無新成績更新")

            time.sleep(60) 
            driver.refresh()

    except Exception as e:
        logging.error(f"執行期間發生錯誤: {e}")
        return "RESTART"
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("瀏覽器已關閉。")
            except:
                pass

# ================= 主程式 (無限迴圈) =================
def main():
    nkust_id, nkust_pwd, target_email = get_credentials()
    logging.info(f"程式啟動，使用者: {nkust_id}")
    print("💡 程式將無限循環執行。若發生登出或錯誤，會自動重啟新視窗。")
    print("💡 若要完全關閉程式，請直接關閉此黑色視窗 (CMD)。")

    while True:
        # 執行任務
        status = run_browser_task(nkust_id, nkust_pwd, target_email)
        
        if status == "RESTART":
            logging.info("⏳ 等待 5 秒後重新啟動系統...")
            time.sleep(5)
            logging.info("🔄 正在重新啟動...")
            continue # 跳回迴圈開頭，重新執行 run_browser_task
        else:
            logging.info("程式意外結束，5 秒後重試...")
            time.sleep(5)

if __name__ == "__main__":
    main()