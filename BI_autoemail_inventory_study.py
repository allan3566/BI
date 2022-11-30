import pyautogui
import time

from datetime import date
import datetime
import win32com.client as client

import os
import shutil
import webbrowser
import sys
class auto_email():

    def __init__(self):
        self

    def open_BI_web(self, address: str, BI_wait_time: int) ->none:
        pyautogui.moveTo(2, 2)
        webbrowser.open(address)
        time.sleep(BI_wait_time)
        pyautogui.press('alt')
        time.sleep(0.25)
        pyautogui.press('left')
        time.sleep(0.25)
        pyautogui.press('left')
        time.sleep(0.25)
        pyautogui.press('left')
        time.sleep(0.25)
        pyautogui.press('left')
        time.sleep(0.25)
        pyautogui.press('left')
        time.sleep(0.25)
        pyautogui.press('enter')
        time.sleep(5)
        pyautogui.hotkey('alt', 'f4')

    def copyCertainFiles(self, source_folder: str, dest_folder: str, string_to_match: str, file_type=None) ->none:

        # Check all files in source_folder
        for filename in os.listdir(source_folder):
            # Move the file if the filename contains the string to match
            if file_type == None:
                if string_to_match in filename:
                    shutil.move(os.path.join(source_folder, filename), dest_folder)
                    os.rename(os.path.join(dest_folder, filename), os.path.join(dest_folder, "first.jpg"))

            # Check if the keyword and the file type both match
            elif isinstance(file_type, str):
                if string_to_match in filename and file_type in filename:
                    shutil.move(os.path.join(source_folder, filename), dest_folder)
                    os.rename(os.path.join(dest_folder, filename), os.path.join(dest_folder, "first.jpg"))


    def writeEmail(self, title: str, sender: str, CC) ->none:
        today = date.today()
        today = today.strftime("%Y/%m/%d")

        # startup and instance of outlook
        outlook = client.Dispatch('Outlook.Application')
        shell = client.Dispatch("WScript.Shell")
        time.sleep(5)

        # create a message
        html_body = '<img src="C:\\Users\\it.ap\\Downloads\\screencaptures\\first.jpg">'\

        message = outlook.CreateItem(0)
        #message.To = 'Allan.Tien@medicom-taiwan.com;mike.li@medicom-taiwan.com;grace.tseng@medicom-taiwan.com;Daphne.Wang@medicom-asia.com'

        # set the message properties
        message.To = sender
        message.CC = CC
        message.Subject = title + today
        message.HTMLBody = html_body

        # display the message to review
        message.Display()
        message.Send()

    def check_outlookSending_deleteFile(self) ->none:
        pyautogui.press('win')
        time.sleep(1)
        pyautogui.write('outlook',interval=0.2)
        time.sleep(1)
        pyautogui.press('enter')
        time.sleep(30)
        pyautogui.moveTo(2, 2)
        pyautogui.click()
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)
        pyautogui.hotkey('alt', 'f4')
        time.sleep(1)

        os.remove(r'C:\Users\it.ap\Downloads\screencaptures\first.jpg')
        time.sleep(1)


if __name__ == '__main__':

    auto_email().open_BI_web(r'https://biserverw2016.medicom.com.hk/reports/powerbi/5600-MTT/Inventory%20study?rs:embed=true',8)
    time.sleep(10)
    auto_email().copyCertainFiles(r'C:\Users\it.ap\Downloads',r'C:\Users\it.ap\Downloads\screencaptures','screencapture-')
    auto_email().writeEmail('SCM_REPORT','gilbert.cheung@medicom-asia.com; dominic.lee@medicom-asia.com; Kurt.Wang@medicom-asia.com','Allan.Tien@medicom-taiwan.com;daphne.wang@medicom-asia.com')
    auto_email().check_outlookSending_deleteFile()