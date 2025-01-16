import pyautogui
import time
import sys
import pyperclip


#Abrir Excel
while True:
    pyautogui.click(1367,1043)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)

# Verificar se a célula está vazia
    clipboard_data = pyperclip.paste()
    if clipboard_data.strip() == '':
        print('Célula vazia detectada. Encerrando o programa.')
        sys.exit()

    #Abrir Protheus
    pyautogui.moveTo(1420,1054)
    pyautogui.click()
    time.sleep(2)
    pyautogui.doubleClick(1572,194)
    pyautogui.hotkey('ctrl', 'v')

    #Click Lupa
    pyautogui.doubleClick(1815, 186)
    time.sleep(2)

    #Click Alterar
    pyautogui.click(183,189)
    time.sleep(9)
    pyautogui.click(1152,613)
    time.sleep(2)

    #Copiar Código Município
    pyautogui.click(1367,1043)
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    #Abrir Protheus
    pyautogui.click(1420,1054)
    time.sleep(2)
    pyautogui.doubleClick(1179,615)
    pyautogui.hotkey('ctrl', 'v')
    pyautogui.press('tab')

    break