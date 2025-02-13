#Bu kod M.Ertuğrul Yıldırım Tarafından MSS Savunma şirketi Malzeme Planlama Departmanı için 13.02.2025 Tarihinde geliştirilmiştir. Hiçbir fidye yazılım kullanmaz. Kodların açıklamaları yazmaktadır. Apache 2.0 Lisansı ile lisanslanmıştır.
import win32com.client as win32
import PyPDF2
import pyautogui
import time
import re
import datetime
import os
import pyperclip
import pygetwindow as gw
import subprocess
import pandas as pd
import xlrd

def check_and_open_window(title_keywords, executable_path):
    if not check_window(title_keywords):
        print("Program çalıştırılıyor...")
        subprocess.Popen([executable_path])
        time.sleep(5)  # Programın açılması için bekleme süresi
    return check_window(title_keywords)

def check_window(title_keywords):
    for window in gw.getAllTitles():
        if any(keyword in window for keyword in title_keywords):
            target_window = gw.getWindowsWithTitle(window)[0]
            if not target_window.isActive:
                print(f"Pencere aktive ediliyor: {window}")
                target_window.activate()
                time.sleep(1)  # Pencerenin aktif hale gelmesi için kısa süre bekleme
            return True
    return False

def save_attachments(save_path):
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6)  # 6 = Inbox
    one_hour_ago = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=1)

    PR_RECEIVED_TIME = "http://schemas.microsoft.com/mapi/proptag/0x0E060040"

    print(f"Kontrol edilen zaman aralığı: {one_hour_ago} ve sonrası")

    for message in inbox.Items:
        received_time = message.PropertyAccessor.GetProperty(PR_RECEIVED_TIME)
        received_time = received_time.astimezone(datetime.timezone.utc)
        print(f"İşlenen mesajın alındığı zaman: {received_time}")

        if received_time >= one_hour_ago:
            if message.Attachments.Count > 0:
                for attachment in message.Attachments:
                    attachment_name = attachment.FileName.lower()
                    if any(keyword in attachment_name for keyword in ["sakarya", "43", "4300"]):
                        attachment_path = f"{save_path}/{attachment.FileName}"
                        attachment.SaveAsFile(attachment_path)
                        print(f"Ek kaydedildi: {attachment_path}")

def find_and_copy_text(file_path):
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            if "4300" in text:
                match = re.search(r'\d{5}-\d{4}-\d{2}', text)
                if match:
                    extracted_text = match.group(0)
                    pyperclip.copy(extracted_text)
                    print(f"Kopyalanan metin: {extracted_text}")
                    return extracted_text
    return None

def perform_clicks(text_to_paste, executable_path):
    if not check_and_open_window(["Ekip Barkod"], executable_path):
        raise Exception("Gerekli pencere bulunamadı ve açılamadı. İşlemler durduruldu.")
    
    pyautogui.click(705, 108)
    time.sleep(1)
    
    pyautogui.write(text_to_paste)
    time.sleep(1)
    
    pyautogui.click(1180, 916)
    time.sleep(1)
    
    pyautogui.write("03.14.25")
    time.sleep(1)
    
    pyautogui.click(1180, 916)
    time.sleep(1)
    
    pyautogui.click(629, 97)
    time.sleep(3)

    pyautogui.write("output")
    time.sleep(1)
   
    pyautogui.press("tab")
    time.sleep(1)
    pyautogui.press("down")
    pyautogui.press("down")
    pyautogui.press("down")
    time.sleep(1)
    pyautogui.press("enter")
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.press("enter")
    time.sleep(2)
    
    file_save_path = "C:\\Users\\Hp\\Desktop\\Ekipbarkd\\output.pdf"
    print(f"Dosya kaydediliyor: {file_save_path}")
  
    print(f"Tıklama işlemi tamamlandı ve metin yapıştırıldı: {text_to_paste} ve 03.14.25")
    

def read_excel_value(excel_path, sheet_name, column_indices, row):
    """Excel dosyasındaki belirtilen sütunlarda verileri kontrol et"""
    if not os.path.exists(excel_path):
        print(f"Excel dosyası bulunamadı: {excel_path}")
        return None  # Dosya bulunamadığında None döndür
    
    # Excel dosyasını oku
    df = pd.read_excel(excel_path, sheet_name=sheet_name, engine='xlrd')
    
    # Sütunlar arasında sırasıyla arama yap
    for column_index in column_indices:
        value = df.iat[row, column_index]  # Satır ve sütun index ile değeri al
        print(f"{column_index+1}. sütundaki {row+1}. satırdaki değer: {value}")
        
        # Eğer sütunda veri varsa, o değeri döndür
        if pd.notna(value):  # NaN olmayan değer
            return value
    return None  # Hiçbir sütunda veri yoksa None döndür

def send_email_via_outlook(subject, body, recipient):
    """Outlook üzerinden e-posta gönderme"""
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)  # 0, MailItem'ı temsil eder
    mail.Subject = subject
    mail.Body = body
    mail.To = recipient
    mail.Send()
    print(f"E-posta başarıyla gönderildi: {recipient}")

def compare_and_send_email(output_value, excel_value, required_quantity, missing_quantity):
    """Değerleri karşılaştır ve e-posta gönder"""
    if output_value == excel_value:
        print("Veriler eşleşti, gerekli işlem yapılıyor...")
        # E-posta gönderme işlemi
        subject = "Bu bir deneme mailidir"
        body = "Murat Bey, Sevki uygundur"
        recipient = "purchasing2@msssavunma.com"
        send_email_via_outlook(subject, body, recipient)
    else:
        print(f"Veriler eşleşmedi! Excel'deki değer: {excel_value}, Output'daki değer: {output_value}")
        # Eksik miktar hesabı
        if missing_quantity > 0:
            print(f"Eksik miktar: {missing_quantity}")
            subject = "Eksik Ürün Bildirimi"
            body = f"Eksik miktar: {missing_quantity}"
            recipient = "purchasing2@msssavunma.com"
            send_email_via_outlook(subject, body, recipient)
        else:
            print("Eksik miktar bulunmamaktadır.")

def main():
    save_path = "C:/Users/Hp/Desktop/Mailekleri"
    executable_path = "C:/Path/To/EkipBarkod.exe"

    if not os.path.exists(save_path):
        os.makedirs(save_path)
    
    required_quantity = None
    missing_quantity = None
    output_value = None
    
    for file_name in os.listdir(save_path):
        if file_name.endswith(".pdf"):
            file_path = os.path.join(save_path, file_name)
            print(f"İşlenen dosya: {file_path}")
            if "İstenilen" in file_name:
                required_quantity = find_and_copy_text(file_path)
            elif "Eksik" in file_name:
                missing_quantity = find_and_copy_text(file_path)
            copied_text = find_and_copy_text(file_path)
            if copied_text:
                perform_clicks(copied_text, executable_path)
            
            # Output dosyasındaki veriyi al
            output_excel_path = "C:/Users/Hp/Desktop/Ekipbarkd/output.xls"  # Burada dosya yolunu doğru yazıyoruz
            # AY (50), AZ (51), BA (52) sütunlarında arama yap
            output_value = read_excel_value(output_excel_path, "Sheet1", [43, 44], 28)  # 30. satır için 29. index

            # Karşılaştırma ve mail gönderme işlemi
            if output_value:
                excel_value = output_value  # Burada veriyi karşılaştırabiliriz
                compare_and_send_email(output_value, excel_value, required_quantity, missing_quantity)

if __name__ == "__main__":
    main()
