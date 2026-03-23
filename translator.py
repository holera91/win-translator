import json
import pyperclip
import time
from deep_translator import GoogleTranslator
import sys
import os
import win32clipboard
import win32con
import win32api
import win32gui
import win32event
import threading
from pathlib import Path

SINGLE_INSTANCE_MUTEX_NAME = 'Global\\WinTranslatorHotkeyMutex'
import tkinter as tk
from tkinter import ttk, messagebox
import pystray
from PIL import Image, ImageDraw

CONFIG_FILE = Path(__file__).with_suffix('.json')

DEFAULT_HOTKEYS = {
    'de': {'mod': 'Ctrl+Shift', 'key': 'g'},
    'en': {'mod': 'Ctrl+Shift', 'key': 'e'},
    'fr': {'mod': 'Ctrl+Shift', 'key': 'f'},
    'es': {'mod': 'Ctrl+Shift', 'key': 's'},
    'it': {'mod': 'Ctrl+Shift', 'key': 'i'},
    'exit': {'mod': 'Ctrl+Alt', 'key': 'q'},
}

MODIFIER_MASKS = {
    'Ctrl+Alt': win32con.MOD_CONTROL | win32con.MOD_ALT,
    'Ctrl+Shift': win32con.MOD_CONTROL | win32con.MOD_SHIFT,
    'Ctrl+Alt+Shift': win32con.MOD_CONTROL | win32con.MOD_ALT | win32con.MOD_SHIFT,
}

LANG_ORDER = ['de', 'en', 'fr', 'es', 'it']
LANG_LABELS = {
    'de': 'Німецька',
    'en': 'Англійська',
    'fr': 'Французька',
    'es': 'Іспанська',
    'it': 'Італійська',
}

hotkeys_config = {}
hotkey_window = None
tray_icon = None
settings_window = None
mutex_handle = None


def load_config():
    global hotkeys_config
    if CONFIG_FILE.exists():
        try:
            with CONFIG_FILE.open('r', encoding='utf-8') as f:
                data = json.load(f)
            # Забезпечуємо, що всі ключі присутні
            for key in DEFAULT_HOTKEYS:
                if key not in data:
                    data[key] = DEFAULT_HOTKEYS[key]
        except Exception:
            data = DEFAULT_HOTKEYS.copy()
    else:
        data = DEFAULT_HOTKEYS.copy()
    hotkeys_config = data
    save_config()


def save_config():
    with CONFIG_FILE.open('w', encoding='utf-8') as f:
        json.dump(hotkeys_config, f, ensure_ascii=False, indent=2)


def get_vk_code(char):
    if not char or len(char) != 1:
        raise ValueError('Невалідна клавіша')
    ch = char.upper()
    if 'A' <= ch <= 'Z' or '0' <= ch <= '9':
        return ord(ch)
    raise ValueError('Потрібна латинська буква або цифра')


def validate_hotkeys(config):
    seen = set()
    for lang in LANG_ORDER:
        item = config.get(lang, {})
        mod = item.get('mod')
        key = item.get('key')
        if not mod or not key:
            raise ValueError(f'Невірна конфігурація для {lang}')
        if mod not in MODIFIER_MASKS:
            raise ValueError(f'Невідомий модифікатор {mod}')
        key_code = get_vk_code(key)
        marker = (MODIFIER_MASKS[mod], key_code)
        if marker in seen:
            raise ValueError('Комбінація клавіш не може повторюватись')
        seen.add(marker)


def get_hotkey_entries():
    entries = []
    for idx, lang in enumerate(LANG_ORDER, start=1):
        item = hotkeys_config.get(lang, DEFAULT_HOTKEYS[lang])
        entries.append((idx, lang, item['mod'], item['key']))
    return entries


def unregister_hotkeys(hwnd):
    try:
        for i in range(1, len(LANG_ORDER) + 2):
            win32gui.UnregisterHotKey(hwnd, i)
    except Exception:
        pass


class HotkeyWindow:
    def __init__(self):
        self.hwnd = None
        self.running = True

    def create_window(self):
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self.wnd_proc
        wc.lpszClassName = 'HotkeyWindow'
        wc.hInstance = win32api.GetModuleHandle(None)
        class_atom = win32gui.RegisterClass(wc)
        self.hwnd = win32gui.CreateWindow(class_atom, 'Hotkey Window', 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None)

    def register_hotkeys(self):
        unregister_hotkeys(self.hwnd)
        for idx, lang, mod, key in get_hotkey_entries():
            try:
                modmask = MODIFIER_MASKS[mod]
                vk = get_vk_code(key)
                win32gui.RegisterHotKey(self.hwnd, idx, modmask, vk)
            except Exception as e:
                pass
        # Вихід Ctrl+Alt+Q для безпеки
        try:
            win32gui.RegisterHotKey(self.hwnd, 99, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q'))
        except:
            pass

    def wnd_proc(self, hwnd, msg, wparam, lparam):
        if msg == win32con.WM_HOTKEY:
            if wparam == 99:
                exit_program()
                return 0
            for idx, lang, mod, key in get_hotkey_entries():
                if wparam == idx:
                    if lang == 'de':
                        threading.Thread(target=on_de, daemon=True).start()
                    elif lang == 'en':
                        threading.Thread(target=on_en, daemon=True).start()
                    elif lang == 'fr':
                        threading.Thread(target=on_fr, daemon=True).start()
                    elif lang == 'es':
                        threading.Thread(target=on_es, daemon=True).start()
                    elif lang == 'it':
                        threading.Thread(target=on_it, daemon=True).start()
                    return 0
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def run(self):
        self.register_hotkeys()
        win32gui.PumpMessages()


def get_clipboard_text():
    try:
        win32clipboard.OpenClipboard()
        try:
            # Спершу перевіряємо Unicode - він найнадійніший для кирилиці
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                return win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
            elif win32clipboard.IsClipboardFormatAvailable(win32con.CF_TEXT):
                return win32clipboard.GetClipboardData(win32con.CF_TEXT).decode('utf-8', errors='ignore')
        finally:
            win32clipboard.CloseClipboard()
    except Exception:
        pass
    return ''


def translate_and_replace(target_lang):
    try:
        from pynput.keyboard import Controller, Key
        controller = Controller()
        
        # Очищаємо затиснуті клавіші модифікатори перед емуляцією Ctrl+C
        time.sleep(0.1)
        controller.release(Key.ctrl)
        controller.release(Key.shift)
        controller.release(Key.alt)
        time.sleep(0.1)

        # Копіюємо текст
        controller.press(Key.ctrl)
        controller.press('c')
        controller.release('c')
        controller.release(Key.ctrl)
        time.sleep(0.4)

        text_to_translate = get_clipboard_text()
        if not text_to_translate or not text_to_translate.strip():
            return

        # Перекладаємо
        translated = GoogleTranslator(source='auto', target=target_lang).translate(text_to_translate)
        
        if not translated:
            return

        # Копіюємо результат назад
        pyperclip.copy(translated)
        time.sleep(0.2)

        # Вставляємо результат
        controller.press(Key.ctrl)
        controller.press('v')
        controller.release('v')
        controller.release(Key.ctrl)
        time.sleep(0.1)

    except Exception:
        pass


def on_de():
    translate_and_replace('de')


def on_en():
    translate_and_replace('en')


def on_fr():
    translate_and_replace('fr')


def on_es():
    translate_and_replace('es')


def on_it():
    translate_and_replace('it')


def open_settings_window():
    try:
        root = tk.Tk()
        root.title('Налаштування гарячих клавіш')
        root.geometry('420x300')
        root.resizable(False, False)
        root.attributes("-topmost", True)

        root.deiconify()
        root.lift()
        root.focus_force()

        rows = {}

        def on_key_entry(event, key_var):
            c = event.keysym
            if len(c) == 1 and c.isprintable():
                key_var.set(c.lower())
            return 'break'

        for idx, lang in enumerate(LANG_ORDER):
            cfg = hotkeys_config.get(lang, DEFAULT_HOTKEYS[lang])
            tk.Label(root, text=LANG_LABELS.get(lang, lang)).grid(row=idx, column=0, padx=10, pady=5, sticky='w')

            mod_var = tk.StringVar(value=cfg['mod'])
            mod_box = ttk.Combobox(root, textvariable=mod_var, state='readonly', width=15)
            mod_box['values'] = list(MODIFIER_MASKS.keys())
            mod_box.grid(row=idx, column=1, padx=10)

            key_var = tk.StringVar(value=cfg['key'])
            key_entry = ttk.Entry(root, textvariable=key_var, width=8)
            key_entry.grid(row=idx, column=2, padx=10)
            key_entry.bind('<Key>', lambda e, kv=key_var: on_key_entry(e, kv))

            rows[lang] = (mod_var, key_var)

        def on_save():
            try:
                new_cfg = {}
                for lang, (mod_var, key_var) in rows.items():
                    mod_val = mod_var.get().strip()
                    key_val = key_var.get().strip().lower()
                    if mod_val not in MODIFIER_MASKS:
                        raise ValueError(f'Невідомий модифікатор для {lang}')
                    if len(key_val) != 1 or not key_val.isascii():
                        raise ValueError(f'Введіть одну латинську букву для {lang}')
                    new_cfg[lang] = {'mod': mod_val, 'key': key_val}
                
                temp = dict(new_cfg)
                temp['exit'] = hotkeys_config.get('exit', DEFAULT_HOTKEYS['exit'])
                validate_hotkeys(temp)
                hotkeys_config.update(new_cfg)
                save_config()
                
                if hotkey_window:
                    hotkey_window.register_hotkeys()
                    
                messagebox.showinfo('Налаштування', 'Гарячі клавіші оновлено')
                root.destroy()
            except Exception as e:
                messagebox.showerror('Помилка', str(e))

        def on_close():
            root.destroy()

        save_btn = ttk.Button(root, text='Зберегти', command=on_save)
        save_btn.grid(row=len(LANG_ORDER), column=0, columnspan=3, pady=20)
        
        root.protocol("WM_DELETE_WINDOW", on_close)
        root.mainloop()
        
    except Exception:
        pass


def create_image():
    width = 64
    height = 64
    image = Image.new('RGB', (width, height), color=(0, 0, 0))
    draw = ImageDraw.Draw(image)
    draw.rectangle([(8, 8), (56, 56)], outline='white', width=4)
    draw.text((16, 18), 'T', fill='white')
    return image


def exit_program():
    global tray_icon, mutex_handle, hotkey_window
    try:
        if tray_icon:
            tray_icon.stop()
        if hotkey_window and hotkey_window.hwnd:
            unregister_hotkeys(hotkey_window.hwnd)
            win32gui.DestroyWindow(hotkey_window.hwnd)
        if mutex_handle:
            try:
                win32event.ReleaseMutex(mutex_handle)
                win32api.CloseHandle(mutex_handle)
            except Exception:
                pass
    except Exception:
        pass
    finally:
        os._exit(0)


def setup_tray():
    global tray_icon
    def on_settings(icon, item):
        threading.Thread(target=open_settings_window, daemon=True).start()

    def on_exit(icon, item):
        exit_program()

    menu = pystray.Menu(
        pystray.MenuItem('Налаштування', on_settings),
        pystray.MenuItem('Вихід', on_exit),
    )
    tray_icon = pystray.Icon('win-translator', create_image(), 'Win Translator', menu)
    tray_icon.run_detached()


def main():
    global mutex_handle
    mutex_handle = win32event.CreateMutex(None, False, SINGLE_INSTANCE_MUTEX_NAME)
    if win32api.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        return

    load_config()
    setup_tray()
    global hotkey_window
    hotkey_window = HotkeyWindow()
    hotkey_window.create_window()
    hotkey_window.run()


if __name__ == '__main__':
    main()
