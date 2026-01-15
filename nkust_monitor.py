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
ENCRYPTED_SMTP_PASS = "cGt2diB4cGVyIG10aGUgb3l3Zg=="  
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587
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

# ================= 新增：主動檢查 Chrome 是否安裝 =================
def check_chrome_installed():
    """檢查電腦是否安裝了 Google Chrome"""
    # 預設的幾個 Chrome 安裝路徑
    paths = [
        os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
    ]
    
    found = any(os.path.exists(p) for p in paths)
    
    # 測試模式：如果你想測試彈窗，可以把這裡改為 found = False
    # found = False 

    if not found:
        msg = "❌ 偵測不到 Google Chrome 瀏覽器！\n\n本程式需要安裝 Chrome 才能運作。\n請前往 Google 官網下載安裝後再重新執行。"
        logging.critical("環境錯誤: 未安裝 Chrome")
        show_alert("環境錯誤", msg)
        os._exit(0) # 徹底強制結束

# =============================================================

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

def get_config():
    """讀取設定檔"""
    config = configparser.ConfigParser()
    if not os.path.exists('config.txt'):
        show_alert("設定檔遺失", "找不到 config.txt 設定檔！")
        os._exit(0)
    try:
        config.read('config.txt', encoding='utf-8')
        return {
            "id": config['User']['Student_ID'],
            "pwd": config['User']['Student_Password'],
            "email": config['Email']['My_Gmail'],
            "app_pw": config['Email']['App_Password']
        }
    except Exception as e:
        show_alert("設定檔錯誤", f"讀取 config.txt 失敗，請確認格式正確。\n錯誤內容: {e}")
        os._exit(0)

# ================= 郵件發送 (使用者自寄自收) =================
def send_grade_update_email(conf, new_grades):
    subject = "【成績通知】有新的成績公布了！"
    rows_html = ""
    for subj, score in new_grades:
        comment, color = "", "black"
        try:
            val = float(score)
            if val < 60: comment, color = "不好意思老師這次撈不動 😭", "red"
            else: comment, color = "恭喜你被老師撈撈上岸了 🎉", "green"
        except: color = "blue"
        rows_html += f"<tr><td style='padding:8px;border:1px solid #ddd;'>{subj}</td><td style='padding:8px;border:1px solid #ddd;color:{color};font-weight:bold;'>{score}</td><td style='padding:8px;border:1px solid #ddd;'>{comment}</td></tr>"

    login_url = "https://stdsys.nkust.edu.tw/student/Account/Login?ReturnUrl=%2Fstudent"
    content = f"<h3>帥哥/美女你好：</h3><p>系統偵測到下列成績：</p><table style='border-collapse: collapse; width: 100%;'>{rows_html}</table><br><p><a href='{login_url}'>點此前往校務系統</a></p>"
    
    msg = MIMEText(content, 'html', 'utf-8')
    msg['Subject'] = Header(subject, 'utf-8')
    msg['From'] = formataddr((Header("高科成績小幫手", 'utf-8').encode(), conf['email']))
    msg['To'] = conf['email']

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(conf['email'], conf['app_pw'])
        server.send_message(msg)
        server.quit()
        logging.info("✅ 成功寄送通知信給自己")
    except Exception as e:
        logging.error(f"❌ 寄信失敗 (請檢查應用程式密碼): {e}")

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
def run_browser_task(conf):
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
                os._exit(0) # 直接結束程式，不要重啟
            else:
                # 其他錯誤則往上拋出，讓外層決定是否重啟
                raise e
        # ========================================================
        
        logging.info("前往登入頁面...")
        driver.get("https://stdsys.nkust.edu.tw/student/Account/Login")
        
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "usernameOrEmailAddress"))
        ).send_keys(conf['id'])
        driver.find_element(By.NAME, "Password").send_keys(conf['pwd'])

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
                        send_grade_update_email(conf, new_updates)
                        save_history(history_grades)
                    else:
                        if check_count % 10 == 0:
                            logging.info(f"[{time.strftime('%H:%M:%S')}] 無新成績更新")

            time.sleep(60) 
            driver.refresh()

    except Exception as e:
        error_msg = str(e)
        # ★★★ 核心修正：精確攔截 Selenium 的 Chrome 遺失錯誤 ★★★
        if "no chrome binary" in error_msg.lower() or "cannot find chrome binary" in error_msg.lower():
            msg = "❌ 偵測不到 Google Chrome 瀏覽器！\n\n本程式需要安裝 Chrome 才能運作。\n請前往 Google 官網下載安裝後再重新執行。"
            logging.critical("環境錯誤: Selenium 找不到 Chrome 執行檔")
            show_alert("環境錯誤", msg)
            os._exit(0)  # 強制殺掉所有進程，防止 main() 的重啟迴圈
            
        logging.error(f"運行錯誤: {e}")
        return "RESTART"
    finally:
        if driver: driver.quit()

# ================= 主程式 =================
def main():
    # 1. 執行前先檢查環境
    check_chrome_installed()
    
    # 2. 獲取帳密
    conf = get_config()
    
    while True:
        if run_browser_task(conf) == "RESTART":
            time.sleep(5)

if __name__ == "__main__":
    main()