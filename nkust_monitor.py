import time
import smtplib
import configparser
import base64
import os
import sys
import logging
import json
import traceback
import ctypes  # 新增：用於彈出視窗
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
ENCRYPTED_SMTP_PASS = ""  
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
SMTP_USER = "3333xixa3333@gmail.com"
SENDER_NAME = "高科成績通知系統"
HISTORY_FILE = "grade_history.json"

# ================= 系統提示工具 (新增) =================
def show_alert(title, message):
    """
    使用 Windows 原生 API 彈出提示視窗
    MB_ICONERROR = 0x10
    """
    try:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)
    except:
        print(f"[{title}] {message}")

def get_credentials():
    config = configparser.ConfigParser()
    if not os.path.exists('config.txt'):
        msg = "找不到 config.txt 設定檔！請確認檔案是否存在。"
        logging.error(msg)
        show_alert("設定檔遺失", msg)
        sys.exit()
    try:
        config.read('config.txt', encoding='utf-8')
        return config['User']['Student_ID'], config['User']['Student_Password'], config['User']['Target_Email']
    except Exception as e:
        msg = f"讀取 config.txt 失敗，格式可能錯誤。\n錯誤訊息: {e}"
        logging.error(msg)
        show_alert("設定檔錯誤", msg)
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

# ================= 郵件發送 =================
def send_grade_update_email(target_email, new_grades):
    subject = "【成績通知】有新的成績公布了！"
    rows_html = ""
    
    for subject_name, score_text in new_grades:
        comment = ""
        score_color = "black"
        try:
            score_val = float(score_text)
            if score_val < 60:
                comment = "不好意思老師這次撈不動 😭"
                score_color = "red"
            else:
                comment = "恭喜你被老師撈撈上岸了 🎉"
                score_color = "green"
        except ValueError:
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
    <p><a href='https://stdsys.nkust.edu.tw/student/Account/Login?ReturnUrl=%2Fstudent'>點此前往校務系統</a></p>
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
    history_grades = load_history()
    
    options = webdriver.ChromeOptions()
    options.add_argument("--window-size=1200,900")
    options.add_argument("--disable-blink-features=AutomationControlled") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--log-level=3")

    driver = None
    try:
        logging.info("🚀 啟動瀏覽器視窗...")
        
        # ========================================================
        # ★★★ 新增：攔截 Chrome 未安裝的錯誤 ★★★
        # ========================================================
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        except WebDriverException as e:
            error_msg = str(e)
            if "cannot find Chrome binary" in error_msg or "binary is not a Chrome executable" in error_msg:
                user_msg = "❌ 偵測不到 Google Chrome 瀏覽器！\n\n本程式需要安裝 Chrome 才能運作。\n請前往 Google 官網下載安裝後再重試。"
                logging.critical("環境錯誤: 未安裝 Chrome")
                show_alert("環境錯誤", user_msg)
                sys.exit() # 直接結束程式，不要重啟
            else:
                # 其他錯誤則往上拋出，讓外層決定是否重啟
                raise e
        # ========================================================
        
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
            if len(driver.find_elements(By.CLASS_NAME, "bi-list")) > 0:
                logging.info("✅ 登入成功！(偵測到選單)")
                break
            
            if time.time() - start_time > timeout_seconds:
                logging.error(f"❌ 登入超時，準備重啟...")
                return "RESTART"

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

            try:
                ok_btns = driver.find_elements(By.CSS_SELECTOR, "button.swal2-confirm")
                if len(ok_btns) > 0 and ok_btns[0].is_displayed():
                    logging.info("✨ 偵測到驗證彈窗，點擊 OK...")
                    ok_btns[0].click()
                    time.sleep(3) 
            except Exception:
                pass

            time.sleep(1)

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
                _ = driver.title 
            except WebDriverException:
                logging.error("⚠️ 瀏覽器視窗似乎被關閉了。準備重啟...")
                return "RESTART"

            if "Login" in driver.current_url:
                logging.warning("⚠️ 偵測到被系統自動登出！")
                logging.warning("🔄 正在準備重新啟動瀏覽器並登入...")
                return "RESTART"

            page_source = driver.page_source

            if "An error occurred while processing your request" in page_source:
                logging.warning("⚠️ 偵測到學校系統錯誤頁面 (Error.)")
                logging.warning("⏳ 等待 10 秒讓伺服器冷靜，將自動刷新...")
                time.sleep(10)
                driver.refresh()
                continue 

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
        # 如果是 sys.exit() 引發的 SystemExit，直接往上拋，不當作錯誤處理
        if isinstance(e, SystemExit):
            raise e
        
        logging.error(f"執行期間發生錯誤: {e}")
        return "RESTART"
    finally:
        if driver:
            try:
                driver.quit()
                logging.info("瀏覽器已關閉。")
            except:
                pass

# ================= 主程式 =================
def main():
    try:
        nkust_id, nkust_pwd, target_email = get_credentials()
        logging.info(f"程式啟動，使用者: {nkust_id}")
        print("💡 程式將無限循環執行。若發生登出或錯誤，會自動重啟新視窗。")
        
        while True:
            status = run_browser_task(nkust_id, nkust_pwd, target_email)
            
            if status == "RESTART":
                logging.info("⏳ 等待 5 秒後重新啟動系統...")
                time.sleep(5)
                logging.info("🔄 正在重新啟動...")
                continue 
            else:
                logging.info("程式意外結束，5 秒後重試...")
                time.sleep(5)
    except SystemExit:
        # 正常退出
        pass
    except Exception as e:
        # 捕捉最外層錯誤，確保視窗不會秒關，讓使用者看到 Log
        logging.critical(f"嚴重錯誤: {e}")
        input("按 Enter 結束...")

if __name__ == "__main__":
    main()