import base64
import time

# 輸入你的應用程式密碼
my_password = "pkvv xper mthe oywf"

# 簡單的混淆加密 (Base64)
encoded_bytes = base64.b64encode(my_password.encode("utf-8"))
encrypted_pwd = encoded_bytes.decode("utf-8")

print("="*30)
print("請將下方這串亂碼複製到主程式中：")
print(encrypted_pwd)
print("="*30)

time.sleep(10)