import requests
import os
import time
from dotenv import load_dotenv

load_dotenv()

class TelegramBot:
    def __init__(self):
        self.token = os.getenv("TELEGRAM_BOT_TOKEN")
        self.chat_id = os.getenv("TELEGRAM_CHAT_ID")
        self.base_url = f"https://api.telegram.org/bot{self.token}"

    def send_message(self, text, retries=3):
        """Send a text message."""
        url = f"{self.base_url}/sendMessage"
        payload = {
            "chat_id": self.chat_id,
            "text": text
        }
        for i in range(retries):
            try:
                response = requests.post(url, json=payload, timeout=10)
                if response.status_code == 200:
                    print("텔레그램 메시지 전송 성공.")
                    return True
                else:
                    print(f"텔레그램 전송 실패 {response.status_code}: {response.text}")
            except Exception as e:
                print(f"텔레그램 연결 예외 발생: {e}")
            
            time.sleep(2 ** i) # Exponential backoff
        return False

    def send_document(self, file_path, caption="", retries=3):
        """문서(CSV 등) 전송."""
        url = f"{self.base_url}/sendDocument"
        
        if not os.path.exists(file_path):
            print(f"파일을 찾을 수 없습니다: {file_path}")
            return False

        for i in range(retries):
            try:
                with open(file_path, "rb") as f:
                    files = {"document": f}
                    data = {"chat_id": self.chat_id, "caption": caption}
                    response = requests.post(url, files=files, data=data, timeout=30)
                    
                    if response.status_code == 200:
                        print(f"텔레그램 파일 전송 성공: {os.path.basename(file_path)}")
                        return True
                    else:
                         print(f"텔레그램 파일 업로드 실패 {response.status_code}: {response.text}")
            except Exception as e:
                 print(f"텔레그램 파일 업로드 예외 발생: {e}")
            
            time.sleep(2 ** i)
        return False
