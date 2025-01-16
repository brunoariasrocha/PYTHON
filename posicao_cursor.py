import pyautogui


while True:
    x, y = pyautogui.position()
    print(f"Posição do cursor: x={x}, y={y}")
