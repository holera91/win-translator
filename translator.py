import pyperclip
import time
from deep_translator import GoogleTranslator
import sys
import win32clipboard
import win32con
import win32api
import win32gui
import threading

# Функція для безпечного отримання даних з буфера обміну
def get_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                return win32clipboard.GetClipboardData(win32con.CF_TEXT).decode('utf-8', errors='ignore')
            elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
        finally:
            win32clipboard.CloseClipboard()
    except Exception as e:
        print(f"  ❌ Помилка отримання буфера: {e}")
    return ""

# Функція для перекладу та заміни тексту
def translate_and_replace(target_lang):
    try:
        from pynput.keyboard import Controller, Key
        controller = Controller()
        
        # Очищаємо буфер перед копіюванням
        old_clipboard = pyperclip.paste()
        
        # Відпускаємо всі клавіші перед початком (важливо для Ctrl+Alt комбінацій)
        time.sleep(0.1)
        controller.release(Key.ctrl)
        controller.release(Key.shift)
        controller.release(Key.alt)
        time.sleep(0.1)
        
        # 1. Копіюємо виділений текст через Ctrl+C
        print("  → Копіюю текст...")
        controller.press(Key.ctrl)
        controller.press('c')
        controller.release('c')
        controller.release(Key.ctrl)
        time.sleep(1.0)  # Більша затримка для обновлення буфера
        
        # Отримуємо текст напряму з буфера обміну Windows
        text_to_translate = get_clipboard_text()
        
        if len(text_to_translate) > 50:
            print(f"  → Скопійовано: '{text_to_translate[:50]}...'")
        else:
            print(f"  → Скопійовано: '{text_to_translate}'")
        sys.stdout.flush()
        
        # Якщо нічого не скопіювали (текст не було виділено)
        if not text_to_translate.strip():
            print("  ⚠️ Текст не був виділений!")
            sys.stdout.flush()
            return

        try:
            # 2. Перекладаємо
            print("  → Перекладаю...")
            translated = GoogleTranslator(source='auto', target=target_lang).translate(text_to_translate)
            
            if len(translated) > 50:
                print(f"  → Результат: '{translated[:50]}...'")
            else:
                print(f"  → Результат: '{translated}'")
            
            # 3. Кладемо переклад у буфер
            pyperclip.copy(translated)
            time.sleep(0.4)
            
            # 4. Вставляємо переклад замість виділеного тексту через Ctrl+V
            print("  → Вставляю...")
            controller.press(Key.ctrl)
            controller.press('v')
            controller.release('v')
            controller.release(Key.ctrl)
            time.sleep(0.8)
            print("  ✅ Готово!")
            sys.stdout.flush()
            
        except Exception as e:
            print(f"  ❌ Помилка перекладу: {e}")
            pyperclip.copy(old_clipboard)
            sys.stdout.flush()
            
    except Exception as e:
        print(f"  ❌ Загальна помилка: {e}")
        sys.stdout.flush()

# Налаштування гарячих клавіш
print("Скрипт перекладу запущено.")
print("Ctrl+Shift+G -> Німецька (German)")
print("Ctrl+Shift+E -> Англійська (English)")
print("Ctrl+Shift+F -> Французька (French)")
print("Ctrl+Shift+S -> Іспанська (Spanish)")
print("Ctrl+Shift+I -> Італійська (Italian)")
print("Натисніть Ctrl+Alt+Q для повного виходу.")
print("Скрипт чекає на гарячі клавіshi...\n")
sys.stdout.flush()

# Гарячі клавіші з debug-повідомленнями
def on_de():
    try:
        print("Ctrl+Shift+G натиснута - переклад на німецьку...")
        sys.stdout.flush()
        translate_and_replace('de')
        print("Переклад завершено!\n")
        sys.stdout.flush()
    except Exception as e:
        print(f"❌ Помилка: {e}\n")
        sys.stdout.flush()

def on_en():
    try:
        print("Ctrl+Shift+E натиснута - переклад на англійську...")
        sys.stdout.flush()
        translate_and_replace('en')
        print("Переклад завершено!\n")
        sys.stdout.flush()
    except Exception as e:
        print(f"❌ Помилка: {e}\n")
        sys.stdout.flush()

def on_fr():
    try:
        print("Ctrl+Shift+F натиснута - переклад на французьку...")
        sys.stdout.flush()
        translate_and_replace('fr')
        print("Переклад завершено!\n")
        sys.stdout.flush()
    except Exception as e:
        print(f"❌ Помилка: {e}\n")
        sys.stdout.flush()

def on_es():
    try:
        print("Ctrl+Shift+S натиснута - переклад на іспанську...")
        sys.stdout.flush()
        translate_and_replace('es')
        print("Переклад завершено!\n")
        sys.stdout.flush()
    except Exception as e:
        print(f"❌ Помилка: {e}\n")
        sys.stdout.flush()

def on_it():
    try:
        print("Ctrl+Shift+I натиснута - переклад на італійську...")
        sys.stdout.flush()
        translate_and_replace('it')
        print("Переклад завершено!\n")
        sys.stdout.flush()
    except Exception as e:
        print(f"❌ Помилка: {e}\n")
        sys.stdout.flush()

# Використовуємо win32api для реєстрації глобальних гарячих клавіш
def register_hotkeys(hwnd):
    # Ctrl+Shift+G (VK_G = 0x47, MOD_CONTROL = 0x02, MOD_SHIFT = 0x04)
    win32gui.RegisterHotKey(hwnd, 1, win32con.MOD_CONTROL | win32con.MOD_SHIFT, 0x47)
    # Ctrl+Shift+E (VK_E = 0x45)
    win32gui.RegisterHotKey(hwnd, 2, win32con.MOD_CONTROL | win32con.MOD_SHIFT, 0x45)
    # Ctrl+Shift+F (VK_F = 0x46)
    win32gui.RegisterHotKey(hwnd, 3, win32con.MOD_CONTROL | win32con.MOD_SHIFT, 0x46)
    # Ctrl+Shift+S (VK_S = 0x53)
    win32gui.RegisterHotKey(hwnd, 4, win32con.MOD_CONTROL | win32con.MOD_SHIFT, 0x53)
    # Ctrl+Shift+I (VK_I = 0x49)
    win32gui.RegisterHotKey(hwnd, 5, win32con.MOD_CONTROL | win32con.MOD_SHIFT, 0x49)
    # Ctrl+Alt+Q (VK_Q = 0x51, MOD_ALT = 0x01)
    win32gui.RegisterHotKey(hwnd, 6, win32con.MOD_CONTROL | win32con.MOD_ALT, 0x51)

def unregister_hotkeys(hwnd):
    try:
        for i in range(1, 7):
            win32gui.UnregisterHotKey(hwnd, i)
    except:
        pass

# Створюємо невидиме вікно для обробки повідомлень
class HotkeyWindow:
    def __init__(self):
        self.hwnd = None
        self.running = True
        
    def create_window(self):
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self.wnd_proc
        wc.lpszClassName = "HotkeyWindow"
        wc.hInstance = win32api.GetModuleHandle(None)
        
        class_atom = win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow(class_atom, "Hotkey Window", 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None)
        
    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == 1:  # Ctrl+Shift+G
                threading.Thread(target=on_de, daemon=True).start()
            elif wparam == 2:  # Ctrl+Shift+E
                threading.Thread(target=on_en, daemon=True).start()
            elif wparam == 3:  # Ctrl+Alt+F
                threading.Thread(target=on_fr, daemon=True).start()
            elif wparam == 4:  # Ctrl+Alt+S
                threading.Thread(target=on_es, daemon=True).start()
            elif wparam == 5:  # Ctrl+Alt+I
                threading.Thread(target=on_it, daemon=True).start()
            elif wparam == 6:  # Ctrl+Alt+Q
                print("Вихід зі скрипту...")
                self.running = False
                win32gui.PostQuitMessage(0)
            return 0  # Важливо повернути 0 для оброблених повідомлень
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)
    
    def run(self):
        # Реєструємо гарячі клавіші після створення вікна
        register_hotkeys(self.hwnd)
        print("Гарячі клавіші зареєстровано (win32gui)\n")
        sys.stdout.flush()
        
        # Використовуємо PumpMessages для обробки повідомлень
        win32gui.PumpMessages()

# Запускаємо обробку гарячих клавіш
try:
    window = HotkeyWindow()
    window.create_window()
    window.run()
    
except Exception as e:
    print(f"❌ Помилка реєстрації гарячих клавіш: {e}")
    sys.exit(1)

finally:
    try:
        if 'window' in locals() and window.hwnd:
            unregister_hotkeys(window.hwnd)
    except:
        pass
    print("Скрипт завершив роботу.")
