import os
import pyautogui
import time

def fechar_apps():
    repetir = 3
    for i in range(repetir):
        pyautogui.hotkey('alt', 'f4')


# Executa o login no tiny para não dar erro na automação
os.startfile('C:\\Users\\User\\Desktop\\autoTinyAmbro - Power Automate.url')
time.sleep(60)

# Abre a automação
os.startfile('C:\\Users\\User\\Desktop\\autoAmbro - Power Automate.url')
time.sleep(2000)
fechar_apps()