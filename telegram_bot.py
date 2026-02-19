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
                    print("Telegram message sent.")
                    return True
                else:
                    print(f"Telegram error {response.status_code}: {response.text}")
            except Exception as e:
                print(f"Telegram exception: {e}")
            
            time.sleep(2 ** i) # Exponential backoff
        return False

    def send_document(self, file_path, caption="", retries=3):
        """Send a document (CSV) with caption."""
        url = f"{self.base_url}/sendDocument"
        
        if not os.path.exists(file_path):
            print(f"File not found: {file_path}")
            return False

        for i in range(retries):
            try:
                with open(file_path, "rb") as f:
                    files = {"document": f}
                    data = {"chat_id": self.chat_id, "caption": caption}
                    response = requests.post(url, files=files, data=data, timeout=30)
                    
                    if response.status_code == 200:
                        print(f"Telegram document sent: {os.path.basename(file_path)}")
                        return True
                    else:
                         print(f"Telegram upload error {response.status_code}: {response.text}")
            except Exception as e:
                 print(f"Telegram upload exception: {e}")
            
            time.sleep(2 ** i)
        return False
