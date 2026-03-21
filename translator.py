import keyboard
import pyperclip
import time
from deep_translator import GoogleTranslator
import pyautogui
import sys

# Функція для перекладу та заміни тексту
def translate_and_replace(target_lang):
    # Очищаємо буфер перед копіюванням, щоб знати, чи успішно скопійовано новий текст
    old_clipboard = pyperclip.paste()
    pyperclip.copy("") 
    
    # 1. Копіюємо виділений текст (імітуємо Ctrl+C)
    # Ми використовуємо pyautogui, оскільки keyboard.press_and_release('ctrl+c') іноді дає збої в RDP
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(0.3)  # Невелика пауза для стабільності буфера (важливо для RDP)
    
    text_to_translate = pyperclip.paste()
    
    # Якщо нічого не скопіювали (текст не було виділено), повертаємо старий буфер і виходимо
    if not text_to_translate.strip():
        pyperclip.copy(old_clipboard)
        return

    try:
        # 2. Перекладаємо (автовизначення мови -> цільова мова)
        translated = GoogleTranslator(source='auto', target=target_lang).translate(text_to_translate)
        
        # 3. Кладемо переклад у буфер
        pyperclip.copy(translated)
        time.sleep(0.1)
        
        # 4. Вставляємо переклад замість виділеного тексту (імітуємо Ctrl+V)
        pyautogui.hotkey('ctrl', 'v')
        
    except Exception as e:
        # У разі помилки просто повертаємо старий текст у буфер
        print(f"Помилка: {e}")
        pyperclip.copy(old_clipboard)

# Налаштування гарячих клавіш
print("Скрипт перекладу запущено.")
print("Ctrl+Shift+G -> Німецька (German)")
print("Ctrl+Shift+E -> Англійська (English)")
print("Натисніть Ctrl+Alt+Q для повного виходу.")

# Гарячі клавіші
keyboard.add_hotkey('ctrl+shift+g', lambda: translate_and_replace('de'))
keyboard.add_hotkey('ctrl+shift+e', lambda: translate_and_replace('en'))

# Тримаємо скрипт у фоні
keyboard.wait('ctrl+alt+q')
