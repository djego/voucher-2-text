import os
import re
import json
import pytesseract
from PIL import Image
from excel import write_to_excel

def process_voucher_to_text():
    voucher_path = "vouchers"
    result = []
    for filename in os.listdir(voucher_path):
        if filename.endswith(".JPEG") or filename.endswith(".JPG"):
            file_path = os.path.join(voucher_path, filename)
            image = Image.open(file_path)        
            text = pytesseract.image_to_string(image)
            if 'BCP' in text:
                print(text)
                amount = re.findall("[\d,]+\.\d+", text)[0]
                operation_number = re.findall(r"\b\d{8,}\b", text)[0]
                date = re.search(r"\b\d+\s\w+\s\d{4}\b", text).group()
                result.append({
                    "bank": 'BCP', 
                    "currency": 'PEN', 
                    "amount": float(amount.replace(",", "")),
                    "number": operation_number, 
                    "date": date}) 
            elif 'Interbank' in text:
                print(text)
                values = {"bank": 'IBK'}
                values["currency"] = 'USD' if "US$" in text else 'PEN'
                if values["currency"] == 'USD':
                    amount = re.search(r"\$\s(\d+\.\d+)", text).group(1)
                    values["amount"] = float(amount.replace(",",""))
                else:
                    amount = re.findall("[\d,]+\.\d+", text)[0]
                    values["amount"] = float(amount.replace(",", ""))
                values["number"] = re.findall(r"\b\d{7,}\b", text)[0]
                values["date"] = re.search(r"Fecha:\s(\d+\s\w+\s\d+)", text).group(1)
                result.append(values)
    return result

if __name__ == '__main__':
    text = process_voucher_to_text()
    write_to_excel(text)