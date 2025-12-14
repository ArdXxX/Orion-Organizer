import customtkinter as ctk
from tkinter import filedialog, messagebox
import pygetwindow as gw
import json
import os
import subprocess
import psutil
import threading
import time
import win32gui
import win32process
import win32con
from datetime import datetime
from queue import Queue

PERFILES_DIR = "perfiles"
LANG_DIR = "languages"
os.makedirs(PERFILES_DIR, exist_ok=True)
os.makedirs(LANG_DIR, exist_ok=True)

perfil_actual = "Default"
monitoreando = False
monitor_thread = None
orion_path = ""

# cola de lanzamientos para aplicar delay entre clientes
launch_queue = Queue()

app = None
console = None
btn_monitoreo = None
btn_restaurar = None
btn_frente = None
btn_guardar_pos = None
btn_config_global = None

btn_add = None
btn_quitar = None
btn_config_client = None
btn_refresh = None

label_available = None
label_saved = None
log_label = None

scroll_disp = None
scroll_guard = None

checkboxes_disp = {}
checkboxes_guard = {}

profile_var = None
profile_menu = None

BTN_WIDTH = 150
BTN_HEIGHT = 36

BASE_LANG = {
    "TITLE_APP": "Orion Organizer",
    "BTN_START_MON": "Start monitoring",
    "BTN_STOP_MON": "Stop monitoring",
    "BTN_RESTORE_POS": "Restore positions",
    "BTN_BRING_FRONT": "Bring to front",
    "BTN_SAVE_POS": "Save current positions",
    "BTN_CONFIG_GLOBAL": "Configuration",
    "LBL_AVAILABLE_CLIENTS": "Available clients",
    "LBL_SAVED_CLIENTS": "Saved clients",
    "BTN_ADD_SELECTED": "Add selected",
    "BTN_REMOVE": "Remove",
    "BTN_CONFIGURE": "Configure client",
    "BTN_REFRESH_CLIENTS": "Refresh clients",
    "LBL_LOG": "Log / Events",
    "MSG_SELECT_WINDOW_TITLE": "Selection",
    "MSG_SELECT_WINDOW": "You must select a saved client.",
    "BTN_SELECT_ORION": "Select Orion folder",
    "LBL_PROFILE": "Profile",
    "BTN_NEW_PROFILE": "New profile",
    "BTN_DELETE_PROFILE": "Delete profile",
    "LBL_LANGUAGE": "Language",
    "NEW_PROFILE_TITLE": "New profile",
    "NEW_PROFILE_NAME": "Profile name:",
    "ERR_PROFILE_EXISTS": "A profile with that name already exists.",
    "ERR_PROFILE_EMPTY": "Profile name cannot be empty.",
    "CONFIRM_DELETE_PROFILE_TITLE": "Delete profile",
    "CONFIRM_DELETE_PROFILE": "Are you sure you want to delete profile '{profile}'?",
    "ERR_CANNOT_DELETE_LAST_PROFILE": "You cannot delete the last profile.",
    "ERR_CANNOT_DELETE_DEFAULT": "You cannot delete the Default profile.",
    "BTN_SAVE": "Save",
    "LBL_SCHEDULE_SECTION": "Schedule",
    "LBL_SHUTDOWN_TIME": "Shutdown (HH:MM):",
    "LBL_STARTUP_TIME": "Startup (HH:MM):",
    "CHK_ENABLE_SCHEDULE": "Enable schedule",
    "LBL_ORION_PROFILE_SECTION": "Orion profile",
    "LBL_ORION_PROFILE_NAME": "Profile name:",
    "LBL_POSITION_SIZE_SECTION": "Position / Size",
    "BTN_USE_CURRENT_POS": "Use current position",
    "LOG_MONITOR_STARTED": "Monitoring started.",
    "LOG_MONITOR_STOPPED": "Monitoring stopped.",
    "LOG_RESTORED_POS": "Position restored for {title}",
    "LOG_NOT_FOUND_RESTORE": "Client not found for restore: {title}",
    "LOG_ADDED_WINDOWS": "{count} clients updated.",
    "LOG_REMOVED_WINDOWS": "{count} clients removed.",
    "LOG_ORION_SET": "Orion path set: {path}",
    "LOG_ORION_LOADED": "Orion path loaded from profile: {path}",
    "LOG_WINDOW_NOT_FOUND_START_CHECK": "[{title}] Not found. Starting double-check…",
    "LOG_WINDOW_WAITING_CONFIRM": "[{title}] Waiting confirmation ({remaining}s)…",
    "LOG_WINDOW_LAUNCHING": "[{title}] Confirmed. Queued for launch…",
    "LOG_WINDOW_LAUNCHING_SCHEDULE": "[{title}] Confirmed (schedule). Queued for launch…",
    "LOG_WINDOW_NO_CMD": "[{title}] No command configured.",
    "LOG_WINDOW_NO_CMD_SCHEDULE": "[{title}] No command configured (schedule).",
    "LOG_WINDOW_ALT_MATCH": "[{title}] Alternative title detected: '{alt}'",
    "LOG_WINDOW_REAPPEARED": "[{title}] Client reappeared. Cancelling double-check.",
    "LOG_SHUTTING_DOWN": "Client '{title}' should be off now.",
    "MENU_NEW_PROFILE": "+ New profile",
    "MENU_DELETE_PROFILE": "Delete profile",
    "MENU_SEPARATOR": "──────────────",
    "LOG_DISCONNECT_DETECTED": "Disconnected client detected: {title}. Closing process...",
    "CFG_TITLE": "Configuration",
    "CFG_TAB_GENERAL": "General",
    "CFG_TAB_TIMINGS": "Timings",
    "CFG_TAB_DISCONNECT": "Disconnect titles",
    "CFG_PROFILE_LABEL": "Active profile:",
    "CFG_THEME_LABEL": "Theme:",
    "CFG_THEME_DARK": "Dark",
    "CFG_THEME_LIGHT": "Light",
    "CFG_MONITOR_INTERVAL": "Monitor interval:",
    "CFG_DCHECK_ENABLED": "Enable double-check",
    "CFG_DCHECK_TIME": "Double-check time:",
    "CFG_REPOSITION_DELAY": "Auto-reposition delay:",
    "CFG_ORION_LABEL": "Orion folder:",
    "CFG_ORION_SELECT": "Select...",
    "CFG_LANG_LABEL": "Language:",
    "CFG_DISCONNECT_LIST_LABEL": "Disconnect titles",
    "CFG_DISCONNECT_ADD": "Add",
    "CFG_DISCONNECT_REMOVE": "Remove selected",
    "CFG_CLOSE": "Close"
}

LANG = BASE_LANG.copy()
LANG_CODE = "en"
LANG_OPTIONS = {}
language_var = None
language_menu = None

def cargar_idiomas_disponibles():
    idiomas = {"en": "English"}
    os.makedirs(LANG_DIR, exist_ok=True)
    for fn in os.listdir(LANG_DIR):
        if not fn.endswith(".json"):
            continue
        code = fn[:-5]
        if code == "en":
            continue
        idiomas[code] = code.capitalize()
    return idiomas

def load_language(code: str):
    global LANG, LANG_CODE
    if code == "en":
        LANG = BASE_LANG.copy()
        LANG_CODE = "en"
        return
    path = os.path.join(LANG_DIR, f"{code}.json")
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(BASE_LANG, f, indent=4, ensure_ascii=False)
        LANG = BASE_LANG.copy()
        LANG_CODE = code
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        merged = BASE_LANG.copy()
        merged.update(data)
        LANG = merged
        LANG_CODE = code
    except Exception:
        LANG = BASE_LANG.copy()
        LANG_CODE = "en"

def tr(key: str, default=None):
    if key in LANG:
        return LANG[key]
    if key in BASE_LANG:
        return BASE_LANG[key]
    return default if default is not None else key

def log(text):
    global console
    if console is None:
        print(text)
        return
    now = datetime.now().strftime("%H:%M:%S")
    console.configure(state="normal")
    console.insert("end", f"[{now}] {text}\n")
    console.see("end")
    console.configure(state="disabled")

def obtener_ruta_config(perfil):
    return os.path.join(PERFILES_DIR, f"{perfil}.json")

def cargar_perfil():
    ruta = obtener_ruta_config(perfil_actual)
    if not os.path.exists(ruta):
        return {
            "orion_path": "",
            "language": "en",
            "windows": [],
            "disconnect_titles": ["Atlantic", "Origin"],
            "double_check_enabled": True,
            "double_check_seconds": 120,
            "auto_reposition_delay": 20,
            "monitor_interval": 10,
            "theme": "dark"
        }
    with open(ruta, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, list):
        data = {
            "orion_path": "",
            "language": "en",
            "windows": data,
            "disconnect_titles": ["Atlantic", "Origin"],
            "double_check_enabled": True,
            "double_check_seconds": 120,
            "auto_reposition_delay": 20,
            "monitor_interval": 10,
            "theme": "dark"
        }
    data.setdefault("orion_path", "")
    data.setdefault("language", "en")
    data.setdefault("windows", [])
    data.setdefault("disconnect_titles", ["Atlantic", "Origin"])
    data.setdefault("double_check_enabled", True)
    data.setdefault("double_check_seconds", 120)
    data.setdefault("auto_reposition_delay", 20)
    data.setdefault("monitor_interval", 10)
    data.setdefault("theme", "dark")
    return data

def guardar_perfil(perfil_data):
    ruta = obtener_ruta_config(perfil_actual)
    with open(ruta, "w", encoding="utf-8") as f:
        json.dump(perfil_data, f, indent=4, ensure_ascii=False)

def cargar_ventanas():
    return cargar_perfil().get("windows", [])

def guardar_ventanas(ventanas):
    perfil_data = cargar_perfil()
    perfil_data["windows"] = ventanas
    guardar_perfil(perfil_data)

def get_orion_path():
    return cargar_perfil().get("orion_path", "")

def set_orion_path(path):
    perfil_data = cargar_perfil()
    perfil_data["orion_path"] = path
    guardar_perfil(perfil_data)

def get_profile_language():
    return cargar_perfil().get("language", "en")

def set_profile_language(code: str):
    perfil_data = cargar_perfil()
    perfil_data["language"] = code
    guardar_perfil(perfil_data)

def get_disconnect_titles():
    return cargar_perfil().get("disconnect_titles", ["Atlantic", "Origin"])

def set_disconnect_titles(titles):
    perfil_data = cargar_perfil()
    perfil_data["disconnect_titles"] = titles
    guardar_perfil(perfil_data)

def listar_perfiles():
    perfiles = []
    for fn in os.listdir(PERFILES_DIR):
        if fn.lower().endswith(".json"):
            perfiles.append(os.path.splitext(fn)[0])
    if not perfiles:
        perfiles = ["Default"]
    return sorted(perfiles)

def ensure_default_profile():
    global perfil_actual
    perfiles = listar_perfiles()
    if "Default" not in perfiles:
        perfil_actual = "Default"
        guardar_perfil({
            "orion_path": "",
            "language": "en",
            "windows": [],
            "disconnect_titles": ["Atlantic", "Origin"],
            "double_check_enabled": True,
            "double_check_seconds": 120,
            "auto_reposition_delay": 20,
            "monitor_interval": 10,
            "theme": "dark"
        })

def actualizar_cmds_orion_en_ventanas():
    global orion_path
    if not orion_path:
        return
    ventanas = cargar_ventanas()
    changed = False
    for v in ventanas:
        perfil = v.get("perfil", "").strip()
        if perfil:
            v["cmd"] = f'pushd "{orion_path}" && OrionLauncher.exe "-autostart|{perfil}"'
            changed = True
    if changed:
        guardar_ventanas(ventanas)

def ordenar_perfiles_para_menu(current, perfiles):
    otros = sorted([p for p in perfiles if p != current])
    if current and current not in otros:
        return [current] + otros
    return otros

def refresh_profile_menu():
    global profile_menu, profile_var
    if profile_menu is None or profile_var is None:
        return
    perfiles = listar_perfiles()
    ordered = ordenar_perfiles_para_menu(perfil_actual, perfiles)
    values = ordered + [tr("MENU_SEPARATOR"), tr("MENU_NEW_PROFILE"), tr("MENU_DELETE_PROFILE")]
    profile_menu.configure(values=values)
    profile_var.set(perfil_actual)

def cambiar_perfil(name: str):
    global perfil_actual, orion_path
    perfil_actual = name
    p = cargar_perfil()
    orion_path = p.get("orion_path", "")
    lang_code = p.get("language", "en")
    if lang_code not in LANG_OPTIONS:
        lang_code = "en"
    load_language(lang_code)
    apply_language()
    actualizar_cmds_orion_en_ventanas()
    actualizar_listas()
    refresh_profile_menu()
    log(f"Profile changed to: {name}")

def on_profile_menu_select(choice: str):
    sep = tr("MENU_SEPARATOR")
    new_lbl = tr("MENU_NEW_PROFILE")
    del_lbl = tr("MENU_DELETE_PROFILE")
    global profile_var
    if profile_var is None:
        return
    if choice == sep:
        profile_var.set(perfil_actual)
        return
    if choice == new_lbl:
        profile_var.set(perfil_actual)
        crear_perfil_nuevo()
        return
    if choice == del_lbl:
        profile_var.set(perfil_actual)
        eliminar_perfil()
        return
    cambiar_perfil(choice)

def crear_perfil_nuevo():
    def do_create():
        name = entry_name.get().strip()
        if not name:
            messagebox.showerror(tr("NEW_PROFILE_TITLE"), tr("ERR_PROFILE_EMPTY"))
            return
        if name in listar_perfiles():
            messagebox.showerror(tr("NEW_PROFILE_TITLE"), tr("ERR_PROFILE_EXISTS"))
            return
        global perfil_actual
        perfil_actual = name
        guardar_perfil({
            "orion_path": "",
            "language": LANG_CODE,
            "windows": [],
            "disconnect_titles": ["Atlantic", "Origin"],
            "double_check_enabled": True,
            "double_check_seconds": 120,
            "auto_reposition_delay": 20,
            "monitor_interval": 10,
            "theme": "dark"
        })
        cambiar_perfil(name)
        refresh_profile_menu()
        win.destroy()

    win = ctk.CTkToplevel(app)
    win.title(tr("NEW_PROFILE_TITLE"))
    win.geometry("320x150")
    win.lift()
    win.focus_force()
    win.grab_set()

    frame = ctk.CTkFrame(win)
    frame.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(frame, text=tr("NEW_PROFILE_NAME")).pack(anchor="w", pady=(5, 5))
    entry_name = ctk.CTkEntry(frame)
    entry_name.pack(fill="x", pady=(0, 10))
    btns = ctk.CTkFrame(frame)
    btns.pack(fill="x")
    ctk.CTkButton(btns, text=tr("BTN_SAVE"), command=do_create).pack(side="right", padx=5, pady=5)

def eliminar_perfil():
    global perfil_actual
    perfiles = listar_perfiles()
    if len(perfiles) <= 1:
        messagebox.showerror(tr("CONFIRM_DELETE_PROFILE_TITLE"), tr("ERR_CANNOT_DELETE_LAST_PROFILE"))
        return
    if perfil_actual == "Default":
        messagebox.showerror(tr("CONFIRM_DELETE_PROFILE_TITLE"), tr("ERR_CANNOT_DELETE_DEFAULT"))
        return
    msg = tr("CONFIRM_DELETE_PROFILE").format(profile=perfil_actual)
    if not messagebox.askyesno(tr("CONFIRM_DELETE_PROFILE_TITLE"), msg):
        return
    ruta = obtener_ruta_config(perfil_actual)
    try:
        if os.path.exists(ruta):
            os.remove(ruta)
    except Exception:
        pass
    restantes = listar_perfiles()
    if not restantes:
        ensure_default_profile()
        restantes = listar_perfiles()
    perfil_actual_nuevo = restantes[0]
    cambiar_perfil(perfil_actual_nuevo)
    refresh_profile_menu()

def parse_hora(hhmm):
    try:
        h, m = map(int, hhmm.split(":"))
        return h * 60 + m
    except Exception:
        return None

def minutos_actuales():
    now = datetime.now()
    return now.hour * 60 + now.minute

def en_franja_apagado(shutdown, startup):
    s_down = parse_hora(shutdown)
    s_up = parse_hora(startup)
    if s_down is None or s_up is None:
        return False
    ahora = minutos_actuales()
    if s_down < s_up:
        return s_down <= ahora < s_up
    else:
        return ahora >= s_down or ahora < s_up

def obtener_clientes_orion():
    clientes = []
    try:
        import re
        for w in gw.getAllWindows():
            if not w.title or not w.visible:
                continue
            title = w.title
            if "- OA" in title:
                continue
            if re.search(r"v\d+\.\d+\.\d+\.\d+", title):
                continue
            try:
                hwnd = w._hWnd
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
                p = psutil.Process(pid)
                if p.name().lower() == "orionuo.exe":
                    clientes.append(w)
            except Exception:
                continue
    except Exception:
        pass
    return clientes

def normalizar_titulo(titulo):
    prefijos = ["Lady ", "Lord "]
    for p in prefijos:
        if titulo.startswith(p):
            return titulo[len(p):]
    return titulo

def buscar_ventana_flexible(titulo):
    wins = gw.getWindowsWithTitle(titulo)
    if wins:
        return True, wins
    titulo_alt = normalizar_titulo(titulo)
    if titulo_alt != titulo:
        wins_alt = gw.getWindowsWithTitle(titulo_alt)
        if wins_alt:
            #log(tr("LOG_WINDOW_ALT_MATCH").format(title=titulo, alt=titulo_alt))
            return True, wins_alt
    return False, []

def cerrar_ventana_por_titulo(titulo):
    existe, wins = buscar_ventana_flexible(titulo)
    if not existe or not wins:
        return False
    win = wins[0]
    hwnd = win._hWnd
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        p = psutil.Process(pid)
        p.terminate()
        log(tr("LOG_SHUTTING_DOWN").format(title=titulo))
        return True
    except Exception as e:
        log(f"Error closing '{titulo}': {e}")
        return False

def detectar_y_terminar_error():
    def enum_windows_callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Error" in title:
                log("Error window detected. Terminating process...")
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    p = psutil.Process(pid)
                    p.terminate()
                    log(f"Process {pid} terminated.")
                except Exception as e:
                    log(f"Error terminating process: {e}")
    try:
        win32gui.EnumWindows(enum_windows_callback, None)
    except Exception as e:
        log(f"Error in error-window detection: {e}")

def es_titulo_desconexion(titulo):
    titles = get_disconnect_titles()
    return titulo.strip() in [t.strip() for t in titles]

def detectar_y_cerrar_desconectados():
    clientes = obtener_clientes_orion()
    for w in clientes:
        t = w.title.strip()
        if es_titulo_desconexion(t):
            log(tr("LOG_DISCONNECT_DETECTED").format(title=t))
            try:
                _, pid = win32process.GetWindowThreadProcessId(w._hWnd)
                p = psutil.Process(pid)
                p.terminate()
                log(f"Process {pid} terminated (disconnect).")
            except Exception as e:
                log(f"Error terminating disconnected client '{t}': {e}")

def restaurar_posiciones():
    ventanas = cargar_ventanas()
    for item in ventanas:
        titulo = item.get("titulo", "")
        pos = item.get("pos")
        if not titulo or not pos or len(pos) != 4:
            continue
        existe, wins = buscar_ventana_flexible(titulo)
        if not existe:
            log(tr("LOG_NOT_FOUND_RESTORE").format(title=titulo))
            continue
        win = wins[0]
        x, y, w, h = pos
        try:
            win.moveTo(x, y)
            win.resizeTo(w, h)
            log(tr("LOG_RESTORED_POS").format(title=titulo))
        except Exception as e:
            log(f"Cannot move/resize {titulo}: {e}")

def traer_al_frente():
    ventanas = cargar_ventanas()
    for item in ventanas:
        titulo = item.get("titulo", "")
        if not titulo:
            continue
        existe, wins = buscar_ventana_flexible(titulo)
        if not existe:
            continue
        win = wins[0]
        try:
            win.minimize()
            win.restore()
            win.activate()
            time.sleep(0.2)
        except Exception as e:
            log(f"Cannot bring to front {titulo}: {e}")

    # Al terminar TODOS → restaurar posiciones
    restaurar_posiciones()


def guardar_posicion_actual():
    ventanas = cargar_ventanas()
    visibles = obtener_clientes_orion()
    cambios = 0
    for item in ventanas:
        titulo = item.get("titulo", "")
        if not titulo:
            continue
        win = next((w for w in visibles if w.title == titulo), None)
        if not win:
            continue
        if win.left < -10000 or win.top < -10000:
            continue
        item["pos"] = [win.left, win.top, win.width, win.height]
        cambios += 1
    guardar_ventanas(ventanas)
    log(tr("LOG_ADDED_WINDOWS").format(count=cambios))

def launch_worker():
    while True:
        cmd, auto_repos_delay = launch_queue.get()
        try:
            subprocess.Popen(cmd, shell=True)
            if auto_repos_delay > 0:
                threading.Timer(auto_repos_delay, restaurar_posiciones).start()
            time.sleep(5)  # delay entre lanzamientos
        except Exception as e:
            log(f"Error launching client: {e}")
        launch_queue.task_done()

def seleccionar_orion_folder():
    global orion_path
    ruta = filedialog.askdirectory(title=tr("BTN_SELECT_ORION"))
    if not ruta:
        return
    if not os.path.exists(os.path.join(ruta, "OrionLauncher.exe")):
        messagebox.showerror("Error", "OrionLauncher.exe not found in selected folder.")
        return
    orion_path = ruta
    set_orion_path(ruta)
    actualizar_cmds_orion_en_ventanas()
    log(tr("LOG_ORION_SET").format(path=ruta))

def monitor_loop():
    global monitoreando
    while monitoreando:
        perfil_cfg = cargar_perfil()
        double_enabled = perfil_cfg.get("double_check_enabled", True)
        double_seconds = perfil_cfg.get("double_check_seconds", 120)
        monitor_interval = perfil_cfg.get("monitor_interval", 10)
        auto_repos_delay = perfil_cfg.get("auto_reposition_delay", 20)

        detectar_y_terminar_error()
        detectar_y_cerrar_desconectados()

        ventanas = cargar_ventanas()
        changed = False

        for item in ventanas:
            titulo = item.get("titulo", "")
            if not titulo:
                continue
            cmd = item.get("cmd", "")
            shutdown = item.get("shutdown", "")
            startup = item.get("startup", "")
            sched = item.get("schedule_enabled", False)
            missing_ts = item.get("last_missing_at")
            existe, _ = buscar_ventana_flexible(titulo)

            if sched and shutdown and startup:
                if en_franja_apagado(shutdown, startup):
                    if existe:
                        cerrar_ventana_por_titulo(titulo)
                    if missing_ts is not None:
                        item["last_missing_at"] = None
                        changed = True
                    continue
                else:
                    if existe:
                        if missing_ts is not None:
                            item["last_missing_at"] = None
                            changed = True
                        continue
                    if not double_enabled:
                        if cmd:
                            log(tr("LOG_WINDOW_LAUNCHING_SCHEDULE").format(title=titulo))
                            launch_queue.put((cmd, auto_repos_delay))
                        else:
                            log(tr("LOG_WINDOW_NO_CMD_SCHEDULE").format(title=titulo))
                        item["last_missing_at"] = None
                        changed = True
                        continue
                    if missing_ts is None:
                        item["last_missing_at"] = time.time()
                        log(tr("LOG_WINDOW_NOT_FOUND_START_CHECK").format(title=titulo))
                        changed = True
                        continue
                    elapsed = time.time() - missing_ts
                    if elapsed < double_seconds:
                        remaining = int(double_seconds - elapsed)
                        log(tr("LOG_WINDOW_WAITING_CONFIRM").format(title=titulo, remaining=remaining))
                        continue
                    if cmd:
                        log(tr("LOG_WINDOW_LAUNCHING_SCHEDULE").format(title=titulo))
                        launch_queue.put((cmd, auto_repos_delay))
                    else:
                        log(tr("LOG_WINDOW_NO_CMD_SCHEDULE").format(title=titulo))
                    item["last_missing_at"] = None
                    changed = True
                    continue

            if existe:
                if missing_ts is not None:
                    item["last_missing_at"] = None
                    changed = True
                continue

            if not double_enabled:
                if cmd:
                    log(tr("LOG_WINDOW_LAUNCHING").format(title=titulo))
                    launch_queue.put((cmd, auto_repos_delay))
                else:
                    log(tr("LOG_WINDOW_NO_CMD").format(title=titulo))
                item["last_missing_at"] = None
                changed = True
                continue

            if missing_ts is None:
                item["last_missing_at"] = time.time()
                log(tr("LOG_WINDOW_NOT_FOUND_START_CHECK").format(title=titulo))
                changed = True
                continue
            elapsed = time.time() - missing_ts
            if elapsed < double_seconds:
                remaining = int(double_seconds - elapsed)
                log(tr("LOG_WINDOW_WAITING_CONFIRM").format(title=titulo, remaining=remaining))
                continue
            if cmd:
                log(tr("LOG_WINDOW_LAUNCHING").format(title=titulo))
                launch_queue.put((cmd, auto_repos_delay))
            else:
                log(tr("LOG_WINDOW_NO_CMD").format(title=titulo))
            item["last_missing_at"] = None
            changed = True

        if changed:
            guardar_ventanas(ventanas)

        time.sleep(monitor_interval)

def toggle_monitoreo():
    global monitoreando, monitor_thread, btn_monitoreo
    if not monitoreando:
        monitoreando = True
        monitor_thread = threading.Thread(target=monitor_loop, daemon=True)
        monitor_thread.start()
        if btn_monitoreo is not None:
            btn_monitoreo.configure(text=tr("BTN_STOP_MON"), fg_color="#8B0000")
        log(tr("LOG_MONITOR_STARTED"))
    else:
        monitoreando = False
        if btn_monitoreo is not None:
            btn_monitoreo.configure(text=tr("BTN_START_MON"), fg_color="green")
        log(tr("LOG_MONITOR_STOPPED"))

def actualizar_listas():
    global checkboxes_disp, checkboxes_guard, scroll_disp, scroll_guard
    for widget in scroll_disp.winfo_children():
        widget.destroy()
    for widget in scroll_guard.winfo_children():
        widget.destroy()
    checkboxes_disp.clear()
    checkboxes_guard.clear()
    clientes_vis = obtener_clientes_orion()
    guardadas = cargar_ventanas()
    tit_guardadas = [v.get("titulo", "") for v in guardadas if v.get("titulo")]
    disponibles = [w for w in clientes_vis if w.title not in tit_guardadas]
    for w in disponibles:
        var = ctk.BooleanVar()
        cb = ctk.CTkCheckBox(scroll_disp, text=w.title, variable=var)
        cb.grid(sticky="w", pady=2)
        checkboxes_disp[w.title] = var
    for v in guardadas:
        titulo = v.get("titulo", "")
        if not titulo:
            continue
        var = ctk.BooleanVar()
        def on_check(vvar=var):
            for t, check in checkboxes_guard.items():
                if check != vvar:
                    check.set(False)
        cb = ctk.CTkCheckBox(scroll_guard, text=titulo, variable=var, command=on_check)
        cb.grid(sticky="w", pady=2)
        checkboxes_guard[titulo] = var

def añadir_seleccionadas():
    clientes_vis = obtener_clientes_orion()
    ventanas = cargar_ventanas()
    añadidos = 0
    for titulo, var in checkboxes_disp.items():
        if not var.get():
            continue
        win = next((w for w in clientes_vis if w.title == titulo), None)
        if not win:
            continue
        if win.left < -10000 or win.top < -10000:
            pos = [0, 0, win.width, win.height]
        else:
            pos = [win.left, win.top, win.width, win.height]
        ventanas.append({
            "titulo": win.title,
            "pos": pos,
            "cmd": "",
            "shutdown": "",
            "startup": "",
            "schedule_enabled": False,
            "perfil": "",
            "last_missing_at": None
        })
        añadidos += 1
    guardar_ventanas(ventanas)
    log(tr("LOG_ADDED_WINDOWS").format(count=añadidos))
    actualizar_listas()

def quitar_seleccionadas():
    ventanas = cargar_ventanas()
    nuevos = []
    quitados = 0
    for item in ventanas:
        titulo = item.get("titulo", "")
        if titulo in checkboxes_guard and checkboxes_guard[titulo].get():
            quitados += 1
        else:
            nuevos.append(item)
    guardar_ventanas(nuevos)
    log(tr("LOG_REMOVED_WINDOWS").format(count=quitados))
    actualizar_listas()

def abrir_configuracion_cliente():
    seleccionados = [t for t, var in checkboxes_guard.items() if var.get()]
    if not seleccionados:
        messagebox.showinfo(tr("MSG_SELECT_WINDOW_TITLE"), tr("MSG_SELECT_WINDOW"))
        return
    titulo = seleccionados[0]
    ventanas = cargar_ventanas()
    ventana = next((d for d in ventanas if d.get("titulo", "") == titulo), None)
    if not ventana:
        return
    top = ctk.CTkToplevel(app)
    top.title(f"{tr('BTN_CONFIGURE')}: {titulo}")
    top.geometry("420x470")
    top.lift()
    top.focus_force()
    top.grab_set()

    frame_hor = ctk.CTkFrame(top)
    frame_hor.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_hor, text=tr("LBL_SCHEDULE_SECTION"),
                 font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
    row1 = ctk.CTkFrame(frame_hor)
    row1.pack(fill="x", pady=5)
    ctk.CTkLabel(row1, text=tr("LBL_SHUTDOWN_TIME"), width=140, anchor="w").pack(side="left")
    entry_shutdown = ctk.CTkEntry(row1, width=100)
    entry_shutdown.pack(side="left", padx=5)
    entry_shutdown.insert(0, ventana.get("shutdown", ""))
    row2 = ctk.CTkFrame(frame_hor)
    row2.pack(fill="x", pady=5)
    ctk.CTkLabel(row2, text=tr("LBL_STARTUP_TIME"), width=140, anchor="w").pack(side="left")
    entry_startup = ctk.CTkEntry(row2, width=100)
    entry_startup.pack(side="left", padx=5)
    entry_startup.insert(0, ventana.get("startup", ""))
    switch_prog = ctk.CTkCheckBox(frame_hor, text=tr("CHK_ENABLE_SCHEDULE"))
    switch_prog.pack(anchor="w", pady=5)
    if ventana.get("schedule_enabled"):
        switch_prog.select()

    frame_perf = ctk.CTkFrame(top)
    frame_perf.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_perf, text=tr("LBL_ORION_PROFILE_SECTION"),
                 font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
    row3 = ctk.CTkFrame(frame_perf)
    row3.pack(fill="x", pady=5)
    ctk.CTkLabel(row3, text=tr("LBL_ORION_PROFILE_NAME"), width=140, anchor="w").pack(side="left")
    entry_perfil = ctk.CTkEntry(row3, width=200)
    entry_perfil.pack(side="left", padx=5)
    entry_perfil.insert(0, ventana.get("perfil", ""))

    frame_pos = ctk.CTkFrame(top)
    frame_pos.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_pos, text=tr("LBL_POSITION_SIZE_SECTION"),
                 font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w")
    row4 = ctk.CTkFrame(frame_pos)
    row4.pack(fill="x", pady=5)
    ctk.CTkLabel(row4, text="X:", width=30, anchor="e").pack(side="left")
    entry_x = ctk.CTkEntry(row4, width=70)
    entry_x.pack(side="left", padx=5)
    ctk.CTkLabel(row4, text="Y:", width=30, anchor="e").pack(side="left")
    entry_y = ctk.CTkEntry(row4, width=70)
    entry_y.pack(side="left", padx=5)
    row5 = ctk.CTkFrame(frame_pos)
    row5.pack(fill="x", pady=5)
    ctk.CTkLabel(row5, text="W:", width=30, anchor="e").pack(side="left")
    entry_w = ctk.CTkEntry(row5, width=70)
    entry_w.pack(side="left", padx=5)
    ctk.CTkLabel(row5, text="H:", width=30, anchor="e").pack(side="left")
    entry_h = ctk.CTkEntry(row5, width=70)
    entry_h.pack(side="left", padx=5)

    pos = ventana.get("pos")
    if pos and len(pos) == 4:
        x, y, w, h = pos
        if x > -10000 and y > -10000:
            entry_x.insert(0, x)
            entry_y.insert(0, y)
            entry_w.insert(0, w)
            entry_h.insert(0, h)

    def usar_actual():
        win = next((w for w in obtener_clientes_orion() if w.title == titulo), None)
        if win:
            entry_x.delete(0, "end"); entry_x.insert(0, win.left)
            entry_y.delete(0, "end"); entry_y.insert(0, win.top)
            entry_w.delete(0, "end"); entry_w.insert(0, win.width)
            entry_h.delete(0, "end"); entry_h.insert(0, win.height)

    def guardar():
        global orion_path
        ventana["shutdown"] = entry_shutdown.get().strip()
        ventana["startup"] = entry_startup.get().strip()
        ventana["schedule_enabled"] = bool(switch_prog.get())
        perfil_v = entry_perfil.get().strip()
        ventana["perfil"] = perfil_v
        if orion_path and perfil_v:
            ventana["cmd"] = f'pushd "{orion_path}" && OrionLauncher.exe "-autostart|{perfil_v}"'
        try:
            x = int(entry_x.get())
            y = int(entry_y.get())
            w = int(entry_w.get())
            h = int(entry_h.get())
            ventana["pos"] = [x, y, w, h]
        except ValueError:
            pass
        if "last_missing_at" not in ventana:
            ventana["last_missing_at"] = None
        guardar_ventanas(ventanas)
        log(f"Saved configuration for: {titulo}")
        top.destroy()
        actualizar_listas()

    frame_btn = ctk.CTkFrame(top)
    frame_btn.pack(fill="x", padx=10, pady=10)
    ctk.CTkButton(frame_btn, text=tr("BTN_USE_CURRENT_POS"), command=usar_actual).pack(side="left", padx=5)
    ctk.CTkButton(frame_btn, text=tr("BTN_SAVE"), command=guardar, fg_color="green").pack(side="right", padx=5)

def on_language_change(choice: str):
    code = None
    for c, name in LANG_OPTIONS.items():
        if name == choice:
            code = c
            break
    if not code:
        return
    load_language(code)
    set_profile_language(code)
    apply_language()

def auto_wrap_buttons():
    app.update_idletasks()
    botones = [btn_monitoreo, btn_guardar_pos, btn_restaurar, btn_frente, btn_config_global]

    for b in botones:
        if not b:
            continue

        width = b.winfo_width()
        if width < BTN_WIDTH * 0.9:  # si aún no está dibujado, saltar
            continue

        text = b.cget("text")
        font = ctk.CTkFont()
        text_width = font.measure(text.replace("\n", " "))

        if text_width > width:
            b.configure(wraplength=width - 8)
        else:
            # que quede en una sola línea
            b.configure(wraplength=0)


def apply_language():
    global app, btn_monitoreo, btn_restaurar, btn_frente, btn_guardar_pos
    global label_available, label_saved, log_label, btn_add, btn_quitar, btn_config_client, btn_refresh, btn_config_global
    if app is not None:
        app.title(tr("TITLE_APP"))
    if btn_monitoreo is not None:
        if monitoreando:
            btn_monitoreo.configure(text=tr("BTN_STOP_MON"))
        else:
            btn_monitoreo.configure(text=tr("BTN_START_MON"))
    if btn_restaurar is not None:
        btn_restaurar.configure(text=tr("BTN_RESTORE_POS"))
    if btn_frente is not None:
        btn_frente.configure(text=tr("BTN_BRING_FRONT"))
    if btn_guardar_pos is not None:
        btn_guardar_pos.configure(text=tr("BTN_SAVE_POS"))
    if btn_config_global is not None:
        btn_config_global.configure(text=tr("BTN_CONFIG_GLOBAL"))
    if label_available is not None:
        label_available.configure(text=tr("LBL_AVAILABLE_CLIENTS"))
    if label_saved is not None:
        label_saved.configure(text=tr("LBL_SAVED_CLIENTS"))
    if log_label is not None:
        log_label.configure(text=tr("LBL_LOG"))
    if btn_add is not None:
        btn_add.configure(text=tr("BTN_ADD_SELECTED"))
    if btn_quitar is not None:
        btn_quitar.configure(text=tr("BTN_REMOVE"))
    if btn_config_client is not None:
        btn_config_client.configure(text=tr("BTN_CONFIGURE"))
    if btn_refresh is not None:
        btn_refresh.configure(text=tr("BTN_REFRESH_CLIENTS"))
    auto_wrap_buttons()
    refresh_profile_menu()

def abrir_configuracion_global():
    global profile_menu, profile_var, language_menu, language_var
    cfg = cargar_perfil()
    win = ctk.CTkToplevel(app)
    win.title(tr("CFG_TITLE"))
    win.geometry("520x460")
    win.lift()
    win.focus_force()
    win.grab_set()

    tabview = ctk.CTkTabview(win)
    tabview.pack(fill="both", expand=True, padx=10, pady=10)
    tab_general = tabview.add(tr("CFG_TAB_GENERAL"))
    tab_timings = tabview.add(tr("CFG_TAB_TIMINGS"))
    tab_disc = tabview.add(tr("CFG_TAB_DISCONNECT"))

    # General tab
    frame_prof = ctk.CTkFrame(tab_general)
    frame_prof.pack(fill="x", padx=10, pady=10)
    frame_prof.grid_columnconfigure(0, weight=0)
    frame_prof.grid_columnconfigure(1, weight=1)
    frame_prof.grid_columnconfigure(2, weight=0)
    frame_prof.grid_columnconfigure(3, weight=0)

    ctk.CTkLabel(frame_prof, text=tr("CFG_PROFILE_LABEL")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    perfiles = listar_perfiles()
    import tkinter as tk
    profile_var = tk.StringVar(value=perfil_actual)
    profile_menu = ctk.CTkOptionMenu(
        frame_prof,
        variable=profile_var,
        values=ordenar_perfiles_para_menu(perfil_actual, perfiles),
        command=on_profile_menu_select,
        width=160
    )
    profile_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    btn_new_prof = ctk.CTkButton(frame_prof, text=tr("BTN_NEW_PROFILE"), width=110, command=crear_perfil_nuevo)
    btn_new_prof.grid(row=0, column=2, padx=5, pady=5, sticky="e")
    btn_del_prof = ctk.CTkButton(frame_prof, text=tr("BTN_DELETE_PROFILE"), width=110, fg_color="#8B0000",
                                 hover_color="#660000", command=eliminar_perfil)
    btn_del_prof.grid(row=0, column=3, padx=5, pady=5, sticky="e")

    frame_lang = ctk.CTkFrame(tab_general)
    frame_lang.pack(fill="x", padx=10, pady=10)
    frame_lang.grid_columnconfigure(0, weight=0)
    frame_lang.grid_columnconfigure(1, weight=1)
    ctk.CTkLabel(frame_lang, text=tr("CFG_LANG_LABEL")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    global language_var, language_menu
    language_var = ctk.StringVar(value=LANG_OPTIONS.get(cfg.get("language", "en"), "English"))
    language_menu = ctk.CTkOptionMenu(
        frame_lang,
        variable=language_var,
        values=list(LANG_OPTIONS.values()),
        command=on_language_change,
        width=160
    )
    language_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    frame_theme = ctk.CTkFrame(tab_general)
    frame_theme.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_theme, text=tr("CFG_THEME_LABEL")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    theme_var = ctk.StringVar(value=cfg.get("theme", "dark"))
    theme_menu = ctk.CTkOptionMenu(
        frame_theme,
        variable=theme_var,
        values=["dark", "light"],
        width=120
    )
    theme_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    frame_orion = ctk.CTkFrame(tab_general)
    frame_orion.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_orion, text=tr("CFG_ORION_LABEL")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    orion_label = ctk.CTkLabel(frame_orion, text=get_orion_path() or "Not set")
    orion_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
    def seleccionar_orion_desde_cfg():
        seleccionar_orion_folder()
        orion_label.configure(text=get_orion_path() or "Not set")
    ctk.CTkButton(frame_orion, text=tr("CFG_ORION_SELECT"), command=seleccionar_orion_desde_cfg).grid(
        row=0, column=2, padx=5, pady=5, sticky="w"
    )

    # Timings tab
    frame_dcheck = ctk.CTkFrame(tab_timings)
    frame_dcheck.pack(fill="x", padx=10, pady=10)
    dcheck_var = ctk.BooleanVar(value=cfg.get("double_check_enabled", True))
    chk_dcheck = ctk.CTkCheckBox(frame_dcheck, text=tr("CFG_DCHECK_ENABLED"), variable=dcheck_var)
    chk_dcheck.grid(row=0, column=0, padx=5, pady=5, sticky="w")

    frame_dtime = ctk.CTkFrame(tab_timings)
    frame_dtime.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_dtime, text=tr("CFG_DCHECK_TIME")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    dtime_map = {"30 s": 30, "60 s": 60, "120 s": 120, "300 s": 300}
    reverse_dtime_map = {v: k for k, v in dtime_map.items()}
    current_dtime = cfg.get("double_check_seconds", 120)
    dtime_var = ctk.StringVar(value=reverse_dtime_map.get(current_dtime, "120 s"))
    dtime_menu = ctk.CTkOptionMenu(frame_dtime, variable=dtime_var, values=list(dtime_map.keys()), width=100)
    dtime_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    frame_mint = ctk.CTkFrame(tab_timings)
    frame_mint.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_mint, text=tr("CFG_MONITOR_INTERVAL")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    mint_map = {"5 s": 5, "10 s": 10, "15 s": 15, "30 s": 30}
    reverse_mint_map = {v: k for k, v in mint_map.items()}
    current_mint = cfg.get("monitor_interval", 10)
    mint_var = ctk.StringVar(value=reverse_mint_map.get(current_mint, "10 s"))
    mint_menu = ctk.CTkOptionMenu(frame_mint, variable=mint_var, values=list(mint_map.keys()), width=100)
    mint_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    frame_repos = ctk.CTkFrame(tab_timings)
    frame_repos.pack(fill="x", padx=10, pady=10)
    ctk.CTkLabel(frame_repos, text=tr("CFG_REPOSITION_DELAY")).grid(row=0, column=0, padx=5, pady=5, sticky="w")
    repos_map = {"5 s": 5, "10 s": 10, "15 s": 15, "20 s": 20, "30 s": 30}
    reverse_repos_map = {v: k for k, v in repos_map.items()}
    current_repos = cfg.get("auto_reposition_delay", 20)
    repos_var = ctk.StringVar(value=reverse_repos_map.get(current_repos, "20 s"))
    repos_menu = ctk.CTkOptionMenu(frame_repos, variable=repos_var, values=list(repos_map.keys()), width=100)
    repos_menu.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    # Disconnect titles tab
    frame_disc_list = ctk.CTkFrame(tab_disc)
    frame_disc_list.pack(fill="both", expand=True, padx=10, pady=10)
    ctk.CTkLabel(frame_disc_list, text=tr("CFG_DISCONNECT_LIST_LABEL"),
                 font=ctk.CTkFont(size=14, weight="bold")).pack(anchor="w", pady=(0, 5))
    disc_listbox = ctk.CTkTextbox(frame_disc_list, height=150)
    disc_listbox.pack(fill="both", expand=True, pady=5)
    disc_listbox.configure(state="normal")
    titles = get_disconnect_titles()
    for t in titles:
        disc_listbox.insert("end", t + "\n")

    frame_disc_btns = ctk.CTkFrame(frame_disc_list)
    frame_disc_btns.pack(fill="x", pady=5)
    entry_disc = ctk.CTkEntry(frame_disc_btns, width=160)
    entry_disc.pack(side="left", padx=5)
    def add_disc():
        name = entry_disc.get().strip()
        if not name:
            return
        disc_listbox.insert("end", name + "\n")
        entry_disc.delete(0, "end")
    def remove_disc():
        try:
            idx = disc_listbox.index("insert linestart")
            line = disc_listbox.get(idx, idx + " lineend")
            all_lines = disc_listbox.get("1.0", "end").splitlines()
            disc_listbox.delete("1.0", "end")
            for l in all_lines:
                if l.strip() and l.strip() != line.strip():
                    disc_listbox.insert("end", l.strip() + "\n")
        except Exception:
            pass
    ctk.CTkButton(frame_disc_btns, text=tr("CFG_DISCONNECT_ADD"), command=add_disc, width=90).pack(side="left", padx=5)
    ctk.CTkButton(frame_disc_btns, text=tr("CFG_DISCONNECT_REMOVE"), command=remove_disc, width=140).pack(side="left", padx=5)

    frame_bottom = ctk.CTkFrame(win)
    frame_bottom.pack(fill="x", padx=10, pady=(0, 10))
    def guardar_cfg():
        cfg_local = cargar_perfil()
        cfg_local["double_check_enabled"] = bool(dcheck_var.get())
        cfg_local["double_check_seconds"] = dtime_map.get(dtime_var.get(), 120)
        cfg_local["monitor_interval"] = mint_map.get(mint_var.get(), 10)
        cfg_local["auto_reposition_delay"] = repos_map.get(repos_var.get(), 20)
        cfg_local["theme"] = theme_var.get()
        lines = disc_listbox.get("1.0", "end").splitlines()
        cfg_local["disconnect_titles"] = [l.strip() for l in lines if l.strip()]
        guardar_perfil(cfg_local)
        load_language(cfg_local.get("language", "en"))
        ctk.set_appearance_mode(cfg_local.get("theme", "dark"))
        apply_language()
        win.destroy()
    ctk.CTkButton(frame_bottom, text=tr("BTN_SAVE"), command=guardar_cfg, fg_color="green").pack(side="right", padx=5, pady=5)
    ctk.CTkButton(frame_bottom, text=tr("CFG_CLOSE"), command=win.destroy).pack(side="right", padx=5, pady=5)

def crear_gui():
    global app, console, btn_monitoreo, btn_restaurar, btn_frente, btn_guardar_pos, btn_config_global
    global scroll_disp, scroll_guard, btn_add, btn_quitar, btn_config_client, btn_refresh
    global label_available, label_saved, log_label, LANG_OPTIONS, orion_path

    ensure_default_profile()
    LANG_OPTIONS = cargar_idiomas_disponibles()
    start_lang = get_profile_language()
    if start_lang not in LANG_OPTIONS:
        start_lang = "en"
    load_language(start_lang)

    theme = cargar_perfil().get("theme", "dark")
    ctk.set_appearance_mode(theme)
    ctk.set_default_color_theme("green")

    app = ctk.CTk()
    app.title(tr("TITLE_APP"))
    app.geometry("1000x620")
    app.minsize(900, 580)
    app.grid_rowconfigure(0, weight=0)
    app.grid_rowconfigure(1, weight=1)
    app.grid_rowconfigure(2, weight=0)
    app.grid_columnconfigure(0, weight=1)

    header = ctk.CTkFrame(app)
    header.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 5))
    for c in range(5):
        header.grid_columnconfigure(c, weight=1)

    btn_monitoreo = ctk.CTkButton(
        header,
        text=tr("BTN_START_MON"),
        fg_color="green",
        height=BTN_HEIGHT,
        width=BTN_WIDTH,
        command=toggle_monitoreo
    )
    btn_monitoreo.grid(row=0, column=0, padx=3, pady=3, sticky="ew")

    btn_guardar_pos = ctk.CTkButton(
        header,
        text=tr("BTN_SAVE_POS"),
        height=BTN_HEIGHT,
        width=BTN_WIDTH,
        command=guardar_posicion_actual
    )
    btn_guardar_pos.grid(row=0, column=1, padx=3, pady=3, sticky="ew")

    btn_restaurar = ctk.CTkButton(
        header,
        text=tr("BTN_RESTORE_POS"),
        height=BTN_HEIGHT,
        width=BTN_WIDTH,
        command=restaurar_posiciones
    )
    btn_restaurar.grid(row=0, column=2, padx=3, pady=3, sticky="ew")

    btn_frente = ctk.CTkButton(
        header,
        text=tr("BTN_BRING_FRONT"),
        height=BTN_HEIGHT,
        width=BTN_WIDTH,
        command=traer_al_frente
    )
    btn_frente.grid(row=0, column=3, padx=3, pady=3, sticky="ew")

    btn_config_global = ctk.CTkButton(
        header,
        text=tr("BTN_CONFIG_GLOBAL"),
        height=BTN_HEIGHT,
        width=BTN_WIDTH,
        command=abrir_configuracion_global
    )
    btn_config_global.grid(row=0, column=4, padx=3, pady=3, sticky="ew")

    content = ctk.CTkFrame(app)
    content.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
    content.grid_rowconfigure(0, weight=1)
    content.grid_rowconfigure(1, weight=0)
    content.grid_columnconfigure((0, 1), weight=1)

    panel_left = ctk.CTkFrame(content)
    panel_left.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=5)
    panel_left.grid_rowconfigure(1, weight=1)
    panel_left.grid_columnconfigure(0, weight=1)

    label_available = ctk.CTkLabel(
        panel_left,
        text=tr("LBL_AVAILABLE_CLIENTS"),
        font=ctk.CTkFont(size=14, weight="bold")
    )
    label_available.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

    scroll_disp = ctk.CTkScrollableFrame(panel_left)
    scroll_disp.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
    scroll_disp.grid_columnconfigure(0, weight=1)

    panel_right = ctk.CTkFrame(content)
    panel_right.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=5)
    panel_right.grid_rowconfigure(1, weight=1)
    panel_right.grid_columnconfigure(0, weight=1)

    label_saved = ctk.CTkLabel(
        panel_right,
        text=tr("LBL_SAVED_CLIENTS"),
        font=ctk.CTkFont(size=14, weight="bold")
    )
    label_saved.grid(row=0, column=0, sticky="w", padx=10, pady=(10, 5))

    scroll_guard = ctk.CTkScrollableFrame(panel_right)
    scroll_guard.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
    scroll_guard.grid_columnconfigure(0, weight=1)

    bottom_buttons = ctk.CTkFrame(content)
    bottom_buttons.grid(row=1, column=0, columnspan=2, sticky="ew", padx=10, pady=(0, 10))
    for c in range(4):
        bottom_buttons.grid_columnconfigure(c, weight=1)

    btn_config_client = ctk.CTkButton(
        bottom_buttons,
        text=tr("BTN_CONFIGURE"),
        height=30,
        command=abrir_configuracion_cliente
    )
    btn_config_client.grid(row=0, column=2, sticky="e")

    btn_add = ctk.CTkButton(
        bottom_buttons,
        text=tr("BTN_ADD_SELECTED"),
        height=30,
        command=añadir_seleccionadas
    )
    btn_add.grid(row=0, column=0, sticky="w")

    btn_refresh = ctk.CTkButton(
        bottom_buttons,
        text=tr("BTN_REFRESH_CLIENTS"),
        height=30,
        command=actualizar_listas
    )
    btn_refresh.grid(row=0, column=1, sticky="w")

    btn_quitar = ctk.CTkButton(
        bottom_buttons,
        text=tr("BTN_REMOVE"),
        height=30,
        fg_color="#AA0000",
        command=quitar_seleccionadas
    )
    btn_quitar.grid(row=0, column=3, sticky="e")

    log_frame = ctk.CTkFrame(app)
    log_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(5, 10))
    log_frame.grid_columnconfigure(0, weight=1)

    global log_label, console
    log_label = ctk.CTkLabel(
        log_frame,
        text=tr("LBL_LOG"),
        font=ctk.CTkFont(size=14, weight="bold")
    )
    log_label.grid(row=0, column=0, sticky="w", padx=10, pady=(8, 2))

    console = ctk.CTkTextbox(log_frame, height=120)
    console.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 8))
    console.configure(state="disabled")

    orion_path_loaded = get_orion_path()
    if orion_path_loaded and os.path.exists(os.path.join(orion_path_loaded, "OrionLauncher.exe")):
        globals()["orion_path"] = orion_path_loaded
        log(tr("LOG_ORION_LOADED").format(path=orion_path_loaded))

    apply_language()
    actualizar_listas()
    log("UI ready.")

    # lanzar worker de cola
    worker_thread = threading.Thread(target=launch_worker, daemon=True)
    worker_thread.start()

    app.mainloop()

if __name__ == "__main__":
    crear_gui()
