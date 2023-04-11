import pyautogui
import pyperclip
import time

class DataFromPicture:
    """
    This class takes contract id and jpg file path as input agruments.
    It will press hot-keys for create a new sheet, rename it, press data from picture excel's feature button using hot-key.
    It will paste the jpg file path to select picture, then wait excel to process for 10 seconds.
    
    """
    def __init__(self, contract_id, jpg_path):
        self.contract_id = contract_id
        self.jpg_path = jpg_path
        
    def _workaround_write(self, text):
        """
        This is a work-around for the bug in pyautogui.write() with non-QWERTY keyboards
        It copies the text to clipboard and pastes it, instead of typing it.
        """
        pyperclip.copy(text)
        pyautogui.hotkey('ctrl', 'v')
        pyperclip.copy('')
    
    def get_data_from_picture(self):
        
        # Create new sheet
        pyautogui.hotkey('shift', 'f11')
        time.sleep(1.5)
        # Rename the new sheet name
        pyautogui.hotkey('alt', 'h', 'o', 'r')
        pyautogui.write(self.contract_id)
        pyautogui.press('enter')
        time.sleep(1.5)
        
        # Press Hot-key for data from picture
        pyautogui.hotkey('alt', 'a', 'f', '1', 'p')
        time.sleep(1.5)
        
        # Paste (write) jpg path to select picture then enter
        self._workaround_write(self.jpg_path)
        pyautogui.press('enter')
        time.sleep(20)
        
        # Stop (sleep) Python working for 20 Seconds for reading picture to text
        pyautogui.press('enter')
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(1.5)
        
    def save_file(self):
        
        # Save excel file 
        pyautogui.hotkey('ctrl', 's')
        time.sleep(3)
        
        
if __name__ == '__main__':
    
    import os
    import pandas as pd
    
    #PICTURE_FOLDER_DIR = r'C:\Users\patcharapol.y\Desktop\Projects\New folder\VS code\OCR\pictures'
    EXCEL_FOLDER_DIR = r'C:\Users\patcharapol.y\Desktop\Projects\New folder\VS code\OCR\excels'
    
    df = pd.read_excel(r'C:\Users\patcharapol.y\Desktop\Projects\New folder\VS code\OCR\info_ocr.xlsx')
    df = df[df['jpg_paths'].notnull()].copy()
    contract_ids = df['contract_id'].astype(str).values
    jpg_paths = df['jpg_paths'].values

    # wait for opening excel 5 seconds
    time.sleep(5)
    for i, (id, path) in enumerate(zip(contract_ids, jpg_paths)):
        auto_excel = DataFromPicture(id, path)
        auto_excel.get_data_from_picture()
        if i % 100 == 0 and i != 0:
            auto_excel.save_file()
    auto_excel.save_file()
    
    
    