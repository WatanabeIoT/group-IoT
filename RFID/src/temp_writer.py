import pyautogui
import pyperclip

PATH_TO_WRITETAG_EXE = r"C:\Users\渡邊研究室IoT\Documents\RFtag\WriteTag\WriteTag.exe"

pyautogui.PAUSE = 1.0
pyautogui.hotkey("win", "r")
pyperclip.copy(PATH_TO_WRITETAG_EXE)
pyautogui.hotkey('ctrl','v')
pyautogui.press("enter")
for _ in range(4):
    pyautogui.press("tab")

# 温度書き込み
# TODO: 温度の受信
pyautogui.write("300")

print("end")