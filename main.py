# -*- coding: utf-8 -*-
"""
Переключатель разрешения основного монитора (тумблер 1920x1080 <-> 800x1080).

Логика состояния:
    - state.flag отсутствует или содержит "1920" -> ставим 1920x1080, пишем в флаг "800";
    - state.flag содержит "800"                  -> ставим 800x1080,  пишем в флаг "1920";
    - флаг повреждён / содержит мусор            -> считаем, что флага нет (режим 1920x1080).

Частота обновления не меняется — берётся текущая (обычно 60 Гц).
Работаем только с основным монитором (дисплей-адаптер с индексом 0).

Сборка:
    pyinstaller --onefile --noconsole --uac-admin --name toggle_resolution main.py
"""

import os
import sys
import traceback
from datetime import datetime

import win32api
import win32con

# --- Константы ---------------------------------------------------------------

# Два состояния тумблера: (ширина, высота)
RES_1920 = (1920, 1080)
RES_800 = (800, 1080)

FLAG_NAME = "state.flag"
LOG_NAME = "toggle.log"

# Символические имена кодов возврата ChangeDisplaySettingsEx
DISP_CHANGE_NAMES = {
    win32con.DISP_CHANGE_SUCCESSFUL: "DISP_CHANGE_SUCCESSFUL",
    win32con.DISP_CHANGE_RESTART: "DISP_CHANGE_RESTART",
    win32con.DISP_CHANGE_FAILED: "DISP_CHANGE_FAILED",
    win32con.DISP_CHANGE_BADMODE: "DISP_CHANGE_BADMODE",
    win32con.DISP_CHANGE_NOTUPDATED: "DISP_CHANGE_NOTUPDATED",
    win32con.DISP_CHANGE_BADFLAGS: "DISP_CHANGE_BADFLAGS",
    win32con.DISP_CHANGE_BADPARAM: "DISP_CHANGE_BADPARAM",
    win32con.DISP_CHANGE_BADDUALVIEW: "DISP_CHANGE_BADDUALVIEW",
}


# --- Утилиты -----------------------------------------------------------------


def get_base_dir():
    """Папка, рядом с которой лежат state.flag и toggle.log.

    Для собранного .exe пути берём относительно sys.executable,
    потому что __file__ внутри onefile-архива указывает на временный каталог.
    Для запуска как .py — относительно самого файла.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def log(message):
    """Дописывает строку с таймстампом в toggle.log (append)."""
    try:
        with open(os.path.join(get_base_dir(), LOG_NAME), "a", encoding="utf-8") as f:
            f.write("[{}] {}\n".format(datetime.now().strftime("%Y-%m-%d %H:%M:%S"), message))
    except OSError:
        # Лог недоступен (нет прав/диск только для чтения) — продолжаем работу,
        # не мешая смене разрешения.
        pass


def show_error(text):
    """MessageBox с ошибкой поверх всех окон (работает и без прав администратора)."""
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(
            None, text, "toggle_resolution — ошибка", 0x10 | 0x1000 | 0x40000  # MB_ICONERROR | MB_TOPMOST | MB_SETFOREGROUND
        )
    except Exception:
        # Даже если MessageBox не показался — ошибка уже записана в лог.
        log("Не удалось показать MessageBox: {}".format(traceback.format_exc()))


def read_flag():
    """Читает state.flag. Возвращает 800 или 1920.

    Любая ошибка чтения/разбора (файла нет, пустой, мусор, не число)
    трактуется как «флага нет» -> начальное состояние 1920.
    """
    path = os.path.join(get_base_dir(), FLAG_NAME)
    try:
        with open(path, "r", encoding="utf-8") as f:
            raw = f.read().strip()
        value = int(raw)
        if value in (800, 1920):
            log("state.flag прочитан: {} — переключаем на {}x1080".format(raw, value))
            return value
        # Число есть, но не 800 и не 1920 — считаем флаг повреждённым.
        log("state.flag содержит неожиданное значение '{}' — трактуем как отсутствие флага".format(raw))
        return 1920
    except FileNotFoundError:
        log("state.flag не найден — начальное состояние 1920")
        return 1920
    except (ValueError, OSError):
        # Повреждённый флаг или недоступный файл — как «нет флага».
        log("state.flag повреждён или недоступен — трактуем как отсутствие флага")
        return 1920


def write_flag(value):
    """Записывает новое состояние тумблера в state.flag."""
    path = os.path.join(get_base_dir(), FLAG_NAME)
    with open(path, "w", encoding="utf-8") as f:
        f.write(str(value))
    log("state.flag обновлён: {}".format(value))


def get_primary_device():
    """Возвращает DeviceName основного адаптера (EnumDisplayDevices(None, 0))."""
    device = win32api.EnumDisplayDevices(None, 0)
    return device.DeviceName


def get_current_mode(device_name):
    """Текущий видеорежим основного монитора (ENUM_CURRENT_SETTINGS)."""
    return win32api.EnumDisplaySettings(device_name, win32con.ENUM_CURRENT_SETTINGS)


def find_listed_mode(device_name, width, height, frequency=None):
    """Ищет режим width x height в списке EnumDisplaySettings.

    Возвращает PyDEVMODE этого режима (как его отдал драйвер) или None.
    Если режимов несколько — prefers режим с частотой frequency.
    """
    found = []
    index = 0
    while True:
        try:
            mode = win32api.EnumDisplaySettings(device_name, index)
        except Exception:
            # pywin32 бросает исключение, когда индексы режимов закончились.
            break
        if mode.PelsWidth == width and mode.PelsHeight == height:
            found.append(mode)
        index += 1
    if not found:
        return None
    if frequency is not None:
        for mode in found:
            if mode.DisplayFrequency == frequency:
                return mode
    return found[0]


def is_mode_available(device_name, width, height):
    """Проверяет, что режим width x height есть в списке EnumDisplaySettings."""
    return find_listed_mode(device_name, width, height) is not None


def apply_resolution(device_name, current_mode, width, height):
    """Применяет разрешение через ChangeDisplaySettingsEx с CDS_UPDATEREGISTRY.

    Возвращает код результата (0 = DISP_CHANGE_SUCCESSFUL).
    Частота обновления сохраняется текущая.
    """
    # Берём PyDEVMODE именно из списка режимов драйвера: модификация
    # «текущего» DEVMODE некоторые драйверы отвергают (DISP_CHANGE_BADMODE)
    # при возврате с низкого разрешения на родное.
    devmode = find_listed_mode(device_name, width, height, current_mode.DisplayFrequency)
    if devmode is None:
        return win32con.DISP_CHANGE_BADMODE
    result = win32api.ChangeDisplaySettingsEx(device_name, devmode, win32con.CDS_UPDATEREGISTRY)
    if result == win32con.DISP_CHANGE_BADMODE:
        # Повторная попытка с CDS_RESET: заставляет драйвер переинициализировать
        # список режимов (помогает при возврате на родное разрешение).
        reset_flags = win32con.CDS_UPDATEREGISTRY | getattr(win32con, "CDS_RESET", 0x40000000)
        result = win32api.ChangeDisplaySettingsEx(device_name, devmode, reset_flags)
    return result


# --- Основная логика ---------------------------------------------------------


def main():
    base_dir = get_base_dir()
    log("=== Запуск: {} ===".format(sys.executable))

    # Проверка прав администратора. При сборке с --uac-admin Windows сам
    # запрашивает elevation; если пользователь откажет — процесс не запустится
    # вовсе. Эта ветка срабатывает, когда .exe/.py запущен обходя UAC.
    try:
        import ctypes

        is_admin = bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        is_admin = False

    if not is_admin:
        log("Запуск без прав администратора (отказ UAC или ручной запуск) — выход")
        show_error(
            "Программе требуются права администратора для смены разрешения экрана.\n"
            "Запустите toggle_resolution.exe и подтвердите запрос контроля учётных записей (UAC)."
        )
        return 1

    # 1. Читаем состояние тумблера из флага.
    state = read_flag()
    if state == 1920:
        target_w, target_h = RES_1920
        next_state = 800
    else:
        target_w, target_h = RES_800
        next_state = 1920

    # 2. Текущий режим основного монитора.
    device_name = get_primary_device()
    current_mode = get_current_mode(device_name)
    current_w = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
    current_h = win32api.GetSystemMetrics(win32con.SM_CYSCREEN)
    log(
        "Основной монитор: {}; текущий режим {}x{} @ {} Гц, {} bpp".format(
            device_name,
            current_w,
            current_h,
            current_mode.DisplayFrequency,
            current_mode.BitsPerPel,
        )
    )

    # 3. Целевой режим должен присутствовать в списке доступных.
    if not is_mode_available(device_name, target_w, target_h):
        log("Режим {}x{} отсутствует в EnumDisplaySettings — флаг не изменён".format(target_w, target_h))
        show_error(
            "Разрешение {}x{} недоступно на основном мониторе.\n"
            "Состояние не изменено (state.flag остался без изменений).".format(target_w, target_h)
        )
        return 1

    # 4. Переключаем разрешение.
    log("Переключение: {}x{} -> {}x{} (частота {} Гц сохраняется)".format(
        current_w, current_h, target_w, target_h, current_mode.DisplayFrequency
    ))
    result = apply_resolution(device_name, current_mode, target_w, target_h)

    if result != win32con.DISP_CHANGE_SUCCESSFUL:
        code_name = DISP_CHANGE_NAMES.get(result, "неизвестный код")
        log("ChangeDisplaySettingsEx вернул ошибку: {} ({}) — флаг не изменён".format(result, code_name))
        show_error(
            "Не удалось сменить разрешение на {}x{}.\n"
            "Код ошибки: {} ({}).\n"
            "Состояние не изменено (state.flag остался без изменений).".format(
                target_w, target_h, result, code_name
            )
        )
        return 1

    # 5. Успех — фиксируем новое состояние тумблера.
    write_flag(next_state)
    log("Разрешение успешно переключено на {}x{}".format(target_w, target_h))
    return 0


if __name__ == "__main__":
    exit_code = 1
    try:
        exit_code = main()
    except Exception:
        # Любая непредвиденная ошибка: лог с трассировкой, MessageBox, код 1.
        log("НЕПРЕДВИДЕННАЯ ОШИБКА:\n{}".format(traceback.format_exc()))
        show_error(
            "Произошла непредвиденная ошибка. Подробности — в toggle.log рядом с программой."
        )
        exit_code = 1
    sys.exit(exit_code)
