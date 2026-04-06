import os, json, datetime, sys, subprocess, time
import tkinter as tk
import customtkinter as ctk
import win32gui, win32con, win32ui, win32process
import psutil, pystray, threading, winreg, ctypes, atexit
from PIL import Image, ImageDraw
from ctypes import wintypes


_wsh = None


def resolve_lnk(lnk_path):
    global _wsh
    try:
        if _wsh is None:
            import win32com.client

            _wsh = win32com.client.Dispatch("WScript.Shell")
        sc = _wsh.CreateShortCut(lnk_path)
        target = sc.Targetpath
        return target if target and os.path.exists(target) else None
    except Exception:
        return None


def get_new_data():
    now = datetime.datetime.now()
    return f"{now.strftime('%#I:%M %p')}  •  {now.strftime('%a, %b %#d')}"


_app_index: list = []
_app_index_time: float = 0.0
_APP_INDEX_TTL = 60.0  # rebuild every 60 seconds


def _build_app_index():
    global _app_index, _app_index_time
    search_paths = [
        os.path.join(
            os.environ["ProgramData"], "Microsoft", "Windows", "Start Menu", "Programs"
        ),
        os.path.join(
            os.environ["AppData"], "Microsoft", "Windows", "Start Menu", "Programs"
        ),
    ]
    entries = []
    for base in search_paths:
        if os.path.exists(base):
            for root, _, files in os.walk(base):
                for f in files:
                    if f.endswith(".lnk"):
                        entries.append({"name": f[:-4], "path": os.path.join(root, f)})
    _app_index = entries
    _app_index_time = time.monotonic()


def search_windows_apps(query):
    global _app_index, _app_index_time
    if not _app_index or time.monotonic() - _app_index_time > _APP_INDEX_TTL:
        _build_app_index()
    q = query.lower()
    return [a for a in _app_index if q in a["name"].lower()][:7]


def get_running_apps():
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, param: param.append(hwnd), hwnds)
    valid = []
    blacklist = [
        "Program Manager",
        "Microsoft Text Input Application",
        "NVIDIA GeForce Overlay",
        "Discord Updater",
        "CTkToplevel",
    ]
    for h in hwnds:
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if not title or any(name in title for name in blacklist):
                continue
            style = win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE)
            is_tool = style & win32con.WS_EX_TOOLWINDOW
            is_app = style & win32con.WS_EX_APPWINDOW
            if is_app and not is_tool:
                valid.append({"title": title, "hwnd": h})
            elif not is_tool and win32gui.GetParent(h) == 0:
                valid.append({"title": title, "hwnd": h})
    return valid


def get_exe_from_hwnd(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(pid).exe()
    except Exception:
        return None


def get_icon_from_exe(exe_path):
    try:
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        if not large:
            return None
        for s in small:
            win32gui.DestroyIcon(s)
        hicon = large[0]
        for i in range(1, len(large)):
            win32gui.DestroyIcon(large[i])
        hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc, 32, 32)
        hdc = hdc.CreateCompatibleDC()
        hdc.SelectObject(hbmp)
        hdc.DrawIcon((0, 0), hicon)
        bmpstr = hbmp.GetBitmapBits(True)
        img = Image.frombuffer("RGBA", (32, 32), bmpstr, "raw", "BGRA", 0, 1)
        win32gui.DestroyIcon(hicon)
        hdc.DeleteDC()
        win32gui.ReleaseDC(0, hdc.GetHandleOutput())
        return img
    except Exception:
        return None


ABM_NEW = 0x00000000
ABM_REMOVE = 0x00000001
ABM_SETPOS = 0x00000003
ABE_TOP = 1
ABE_BOTTOM = 3
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
GA_ROOT = 2


class APPBARDATA(ctypes.Structure):
    _fields_ = [
        ("cbSize", wintypes.DWORD),
        ("hWnd", wintypes.HWND),
        ("uCallbackMessage", wintypes.UINT),
        ("uEdge", wintypes.UINT),
        ("rc", wintypes.RECT),
        ("lParam", wintypes.LPARAM),
    ]


global_abd = None


def register_appbar(window_id, bar_height, padding_y, edge=ABE_TOP):
    global global_abd
    global_abd = APPBARDATA()
    global_abd.cbSize = ctypes.sizeof(APPBARDATA)
    global_abd.hWnd = window_id
    global_abd.uEdge = edge
    ctypes.windll.shell32.SHAppBarMessage(ABM_NEW, ctypes.byref(global_abd))
    sw = ctypes.windll.user32.GetSystemMetrics(0)
    sh = ctypes.windll.user32.GetSystemMetrics(1)
    if edge == ABE_TOP:
        global_abd.rc.top, global_abd.rc.bottom = 0, bar_height + padding_y + 10
    else:
        global_abd.rc.top, global_abd.rc.bottom = sh - (bar_height + padding_y + 10), sh
    global_abd.rc.left, global_abd.rc.right = 0, sw
    ctypes.windll.shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(global_abd))


def unregister_appbar():
    global global_abd
    if global_abd:
        ctypes.windll.shell32.SHAppBarMessage(ABM_REMOVE, ctypes.byref(global_abd))
        global_abd = None


def set_taskbar_visibility(visible=True):
    hwnd_tray = win32gui.FindWindow("Shell_TrayWnd", None)
    abd = APPBARDATA()
    abd.cbSize = ctypes.sizeof(abd)
    abd.hWnd = hwnd_tray
    abd.lParam = 0 if not visible else 2
    ctypes.windll.shell32.SHAppBarMessage(10, ctypes.byref(abd))
    cmd = win32con.SW_SHOW if visible else win32con.SW_HIDE
    if hwnd_tray:
        win32gui.ShowWindow(hwnd_tray, cmd)
    hwnd_start = win32gui.FindWindow("Button", None)
    if not hwnd_start:
        hwnd_start = win32gui.FindWindowEx(0, 0, "Button", None)
    if hwnd_start:
        win32gui.ShowWindow(hwnd_start, cmd)
    reset_work_area()


def reset_work_area():
    sw = ctypes.windll.user32.GetSystemMetrics(0)
    sh = ctypes.windll.user32.GetSystemMetrics(1)
    rect = wintypes.RECT(0, 0, sw, sh)
    ctypes.windll.user32.SystemParametersInfoW(0x002F, 0, ctypes.byref(rect), 0x02)


def install_to_startup():
    path = os.path.abspath(__file__)
    cmd = f'"{sys.executable.replace("python.exe", "pythonw.exe")}" "{path}"'
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_WRITE,
        )
        winreg.SetValueEx(key, "Limitens_FloatingBar", 0, winreg.REG_SZ, cmd)
        winreg.CloseKey(key)
    except Exception:
        pass


def create_tray_image():
    image = Image.new("RGB", (64, 64), color=(30, 30, 30))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 24, 56, 40), radius=8, fill=(0, 150, 255))
    return image


def setup_system_tray(app_instance):
    def on_quit(icon, item):
        icon.stop()
        app_instance.after(0, app_instance.safe_exit)

    icon = pystray.Icon(
        "FloatingBar",
        create_tray_image(),
        "Floating Bar",
        pystray.Menu(pystray.MenuItem("Quit", on_quit)),
    )
    icon.run()


class FloatingBar(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Limitens_FloatingBar")
        self.overrideredirect(True)
        CK = "#000001"
        self.configure(fg_color=CK)

        self.pill_frame = ctk.CTkFrame(
            self,
            border_width=1,
            border_color="#3a3a3a",
            corner_radius=20,
            bg_color=CK,
        )
        self.pill_frame.pack(fill="both", expand=True)
        self.pill_frame.grid_rowconfigure(0, weight=1)

        self.left_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        self.center_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        self.right_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")

        DEFAULT_CONFIG = {
            "position": "Top",
            "width_percent": 95,
            "height": 45,
            "bg_color": "#1e1e1e",
            "opacity": 0.631,
            "layout": {
                "left": ["start", "apps", "active_window"],
                "center": ["search"],
                "right": ["taskmanager", "sys_monitor", "clock", "tray"],
            },
            "pinned_apps": [],
        }

        if getattr(sys, "frozen", False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        self.config_path = os.path.join(application_path, "config.json")
        if not os.path.exists(self.config_path):
            with open(self.config_path, "w") as f:
                json.dump(DEFAULT_CONFIG, f, indent=4)
        try:
            with open(self.config_path, "r") as f:
                self.config = DEFAULT_CONFIG | json.load(f)
        except Exception:
            self.config = DEFAULT_CONFIG.copy()

        self.pill_frame.configure(fg_color=self.config.get("bg_color"))
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * (float(self.config.get("width_percent", 90)) / 100))
        h = int(self.config.get("height", 50))

        if self.config.get("position", "Top") == "Bottom":
            y_pos = sh - h - 10
        else:
            y_pos = 10

        self.geo_string = f"{w}x{h}+{(sw - w) // 2}+{y_pos}"
        self.geometry(self.geo_string)
        self.wm_attributes(
            "-transparentcolor",
            CK,
            "-topmost",
            True,
            "-alpha",
            float(self.config.get("opacity", 1.0)),
        )

        # ── Design tokens ──────────────────────────────────────────────────────
        _bg = self.config.get("bg_color", "#1e1e1e")
        _bh = max(26, h - 12)  # button height: 33px at default h=45
        _icr = 8  # corner radius, all buttons
        _icf = ("Segoe MDL2 Assets", 13)  # MDL2 icon font, all icon buttons
        _ich = "#c0c0c0"  # icon glyph colour
        _ihv = "#2a2a2a"  # hover background
        self._btn_h = _bh  # stored so add_app_button can use it

        def _ibtn(parent, glyph, cmd, font=_icf):
            return ctk.CTkButton(
                parent,
                text=glyph,
                width=_bh,
                height=_bh,
                fg_color="transparent",
                hover_color=_ihv,
                corner_radius=_icr,
                font=font,
                text_color=_ich,
                command=cmd,
            )

        # ── Widgets ────────────────────────────────────────────────────────────
        # Active window title
        self.active_window_label = ctk.CTkLabel(
            self.pill_frame,
            text="",
            height=_bh,
            font=("Segoe UI Variable", 11),
            text_color="#9a9a9a",
            anchor="w",
        )

        # Clock — Canvas for pixel-exact two-line vertical centering
        _cw, _cx, _cy = 84, 42, _bh // 2
        self.clock_container = tk.Canvas(
            self.pill_frame,
            width=_cw,
            height=_bh,
            bg=_bg,
            highlightthickness=0,
            bd=0,
        )
        self._clock_time_id = self.clock_container.create_text(
            _cx,
            _cy - 8,
            text="",
            fill="white",
            font=("Segoe UI Variable", 11, "bold"),
            anchor="center",
        )
        self._clock_date_id = self.clock_container.create_text(
            _cx,
            _cy + 8,
            text="",
            fill="#7a7a7a",
            font=("Segoe UI Variable", 9),
            anchor="center",
        )

        # System monitor — Canvas, two-line, mirrors clock layout
        _sw, _sx, _sy = 74, 37, _bh // 2
        self.sys_container = tk.Canvas(
            self.pill_frame,
            width=_sw,
            height=_bh,
            bg=_bg,
            highlightthickness=0,
            bd=0,
        )
        self._sys_cpu_id = self.sys_container.create_text(
            _sx,
            _sy - 8,
            text="CPU  0%",
            fill="#9a9a9a",
            font=("Segoe UI Variable", 10),
            anchor="center",
        )
        self._sys_ram_id = self.sys_container.create_text(
            _sx,
            _sy + 8,
            text="RAM  0%",
            fill="#7a7a7a",
            font=("Segoe UI Variable", 10),
            anchor="center",
        )

        # Separator between sys_monitor and clock — inserted by render_bar
        self._info_sep = ctk.CTkFrame(
            self.pill_frame, width=1, height=16, fg_color="#3a3a3a"
        )

        # App taskbar
        self.apps_container = ctk.CTkFrame(
            self.pill_frame, fg_color="transparent", height=_bh
        )
        self.overflow_btn = _ibtn(
            self.apps_container, "\ue712", self.toggle_overflow_menu
        )
        self.overflow_apps = []

        # Icon buttons — all identical sizing/style
        self.start_btn = _ibtn(
            self.pill_frame,
            "\ue782",
            self.toggle_custom_start,
            font=("Segoe MDL2 Assets", 15),
        )
        self.tray_btn = _ibtn(self.pill_frame, "\ue713", self.toggle_control_center)
        self.search_btn = _ibtn(self.pill_frame, "\ue721", self.toggle_search)
        self.taskmanager_btn = _ibtn(
            self.pill_frame,
            "\ue950",
            lambda: (
                subprocess.Popen("start taskmgr", shell=True)
                if not self.edit_mode
                else None
            ),
        )

        # Edit-mode "Done" button — same height as icon buttons
        self.edit_done_btn = ctk.CTkButton(
            self.pill_frame,
            text="Done",
            width=56,
            height=_bh,
            fg_color="#0078D4",
            hover_color="#106EBE",
            corner_radius=_icr,
            font=("Segoe UI Variable", 11, "bold"),
            text_color="white",
            command=self.toggle_edit_mode,
        )

        self.widget_map = {
            "start": self.start_btn,
            "search": self.search_btn,
            "apps": self.apps_container,
            "active_window": self.active_window_label,
            "taskmanager": self.taskmanager_btn,
            "sys_monitor": self.sys_container,
            "clock": self.clock_container,
            "tray": self.tray_btn,
        }

        self.edit_mode = False
        self.drag_data = {"widget_name": None}
        self.bind_all("<Control-space>", lambda e: self.toggle_search())
        self.bind("<Control-e>", self.toggle_edit_mode)

        for name, widget in self.widget_map.items():
            widget.bind(
                "<ButtonPress-1>", lambda e, n=name: self._on_drag_start(e, n), add="+"
            )
            widget.bind(
                "<B1-Motion>", lambda e, n=name: self._on_drag_motion(e, n), add="+"
            )
            widget.bind(
                "<ButtonRelease-1>", lambda e, n=name: self._on_drag_drop(e, n), add="+"
            )

        self._my_hwnd = None  # cached once after window creation
        self._pending_icons: set = set()  # hwnds currently being loaded in background
        self._prev_hwnds: set = set()  # last known set of open-app hwnds
        self._last_real_title: str = ""  # last focused title from a non-WinBar window
        self._active_hwnd = None  # hwnd of the currently highlighted app button

        psutil.cpu_percent(interval=None)
        # Pre-build the app search index off the main thread so first search is instant
        threading.Thread(target=_build_app_index, daemon=True).start()
        self.render_bar()
        self.check_fullscreen()

    def render_bar(self):
        for zone in [self.left_wing, self.center_wing, self.right_wing]:
            for child in zone.winfo_children():
                child.pack_forget()
        self._info_sep.pack_forget()

        self.pill_frame.grid_columnconfigure(0, weight=1, uniform="wing")
        self.pill_frame.grid_columnconfigure(1, weight=0)
        self.pill_frame.grid_columnconfigure(2, weight=1, uniform="wing")
        self.left_wing.grid(row=0, column=0, sticky="w", padx=(10, 0))
        self.center_wing.grid(row=0, column=1)
        self.right_wing.grid(row=0, column=2, sticky="e", padx=(0, 10))

        layout = self.config.get(
            "layout",
            {
                "left": ["start", "apps", "active_window"],
                "center": ["search"],
                "right": ["taskmanager", "sys_monitor", "clock", "tray"],
            },
        )
        if isinstance(layout, list):
            layout = {"left": layout, "center": [], "right": []}
        for key in layout.get("left", []):
            self.add_widget_to_zone(key, self.left_wing)
        for key in layout.get("center", []):
            self.add_widget_to_zone(key, self.center_wing)
        right_keys = layout.get("right", [])
        for i, key in enumerate(right_keys):
            self.add_widget_to_zone(key, self.right_wing)
            # Insert a thin separator between sys_monitor and clock when adjacent
            if (
                key == "sys_monitor"
                and i + 1 < len(right_keys)
                and right_keys[i + 1] == "clock"
            ):
                self._info_sep.pack(in_=self.right_wing, side="left", padx=2)

    def add_widget_to_zone(self, key, zone):
        if key in self.widget_map:
            self.widget_map[key].pack(in_=zone, side="left", padx=4)

    def _is_own_window(self, hwnd: int) -> bool:
        """Return True if hwnd belongs to WinBar itself or any of its popups."""
        get_ancestor = ctypes.windll.user32.GetAncestor
        if self._my_hwnd and hwnd == self._my_hwnd:
            return True
        for attr in ("start_menu", "search_window", "tray_menu", "overflow_menu"):
            popup = getattr(self, attr, None)
            if popup and popup.winfo_exists():
                if hwnd == get_ancestor(popup.winfo_id(), GA_ROOT):
                    return True
        return False

    def update_time(self):
        now = datetime.datetime.now()
        cpu = int(psutil.cpu_percent())
        ram = int(psutil.virtual_memory().percent)
        self.clock_container.itemconfig(
            self._clock_time_id, text=now.strftime("%#I:%M %p")
        )
        self.clock_container.itemconfig(
            self._clock_date_id, text=now.strftime("%a, %b %#d")
        )
        self.sys_container.itemconfig(self._sys_cpu_id, text=f"CPU {cpu:3}%")
        self.sys_container.itemconfig(self._sys_ram_id, text=f"RAM {ram:3}%")
        self.after(1000, self.update_time)

    def _poll_foreground(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if hwnd and not self._is_own_window(hwnd):
                raw = win32gui.GetWindowText(hwnd)
                self._last_real_title = raw if raw else ""
                self._update_active_indicator(hwnd)
                title = self._last_real_title
                if title:
                    trimmed = title if len(title) <= 30 else title[:29] + "\u2026"
                    self.active_window_label.configure(text=f"\ue76c  {trimmed}")
                else:
                    self.active_window_label.configure(text="")
            elif not hwnd:
                self._update_active_indicator(0)
                self.active_window_label.configure(text="")
        except Exception:
            pass
        self.after(100, self._poll_foreground)

    def _update_active_indicator(self, hwnd: int):
        if hwnd == self._active_hwnd:
            return
        # Clear old indicator — restore to the running-app resting state
        if self._active_hwnd is not None:
            old_btn = self.active_app_buttons.get(self._active_hwnd)
            if old_btn and old_btn.winfo_exists():
                old_btn.configure(fg_color="#252525", border_width=0)
        self._active_hwnd = hwnd
        # Highlight active app with accent border + tinted background
        new_btn = self.active_app_buttons.get(hwnd)
        if new_btn and new_btn.winfo_exists():
            new_btn.configure(
                fg_color="#1a3a52", border_width=1, border_color="#0078D4"
            )

    def _add_app_button(self, hid, title, bar_h, img):
        """Called on the main thread after background icon load completes."""
        self._pending_icons.discard(hid)
        if hid in self.active_app_buttons:
            return
        bh = self._btn_h
        icon_size = int(bh * 0.55)
        icon = ctk.CTkImage(img, img, size=(icon_size, icon_size)) if img else None
        # Running apps: subtle pill bg — visually distinct from transparent pinned launchers
        btn = ctk.CTkButton(
            self.apps_container,
            text="" if icon else title[:8],
            image=icon,
            width=bh,
            height=bh,
            fg_color="#252525",
            hover_color="#303030",
            corner_radius=8,
            command=lambda h=hid: self.focus_window(h),
        )
        btn._full_name = title
        btn._is_pinned = False
        btn.bind("<Button-2>", lambda e, h=hid: self._close_window(h))
        btn.bind("<Enter>", lambda e, w=btn, t=title: self._schedule_tooltip(w, t))
        btn.bind("<Leave>", lambda e: self._hide_tooltip())
        btn.pack(side="left", padx=2)
        self.active_app_buttons[hid] = btn
        # If this window is already the foreground window, highlight it immediately
        if hid == self._active_hwnd:
            btn.configure(fg_color="#1a3a52", border_width=1, border_color="#0078D4")

    def update_open_apps(self):
        if not hasattr(self, "active_app_buttons"):
            self.active_app_buttons = {}
            bh = self._btn_h
            icon_size = int(bh * 0.55)
            pinned = self.config.get("pinned_apps", [])
            for app in pinned:
                if os.path.exists(app["path"]):
                    img = get_icon_from_exe(app["path"])
                    icon = (
                        ctk.CTkImage(img, img, size=(icon_size, icon_size))
                        if img
                        else None
                    )
                    # Pinned: transparent bg — clean launcher style
                    p_btn = ctk.CTkButton(
                        self.apps_container,
                        text="" if icon else app["name"][:3],
                        image=icon,
                        width=bh,
                        height=bh,
                        fg_color="transparent",
                        hover_color="#2a2a2a",
                        corner_radius=8,
                        command=lambda p=app["path"]: subprocess.Popen([p]),
                    )
                    p_btn._full_name = app["name"]
                    p_btn._is_pinned = True
                    p_btn.bind(
                        "<Enter>",
                        lambda e, w=p_btn, t=app["name"]: self._schedule_tooltip(w, t),
                    )
                    p_btn.bind("<Leave>", lambda e: self._hide_tooltip())
                    p_btn.pack(side="left", padx=2)
            if pinned:
                # Visible separator between pinned launchers and live windows
                self._pinned_separator = ctk.CTkFrame(
                    self.apps_container, width=1, height=20, fg_color="#505050"
                )
                self._pinned_separator.pack(side="left", padx=7)

        bar_h = int(self.config.get("height", 50))
        current_windows = get_running_apps()
        cur_hwnds = {app["hwnd"] for app in current_windows}

        # Skip expensive repaint if the window set hasn't changed
        if cur_hwnds == self._prev_hwnds:
            self.after(500, self.update_open_apps)
            return
        self._prev_hwnds = cur_hwnds

        _skip = {"FloatingBar", "Limitens", "NVIDIA", "Overlay", "CTkToplevel"}
        for app in current_windows:
            if any(x in app["title"] for x in _skip):
                continue
            hid = app["hwnd"]
            if hid not in self.active_app_buttons and hid not in self._pending_icons:
                self._pending_icons.add(hid)
                title = app["title"]

                def _load(h=hid, t=title, bh=bar_h):
                    img = get_icon_from_exe(get_exe_from_hwnd(h))
                    self.after(0, lambda: self._add_app_button(h, t, bh, img))

                threading.Thread(target=_load, daemon=True).start()

        for old_id in list(self.active_app_buttons.keys()):
            if old_id not in cur_hwnds:
                if self.active_app_buttons[old_id].winfo_exists():
                    self.active_app_buttons[old_id].destroy()
                del self.active_app_buttons[old_id]

        MAX_VISIBLE = 14
        visible_count = 0
        self.overflow_apps.clear()

        children = self.apps_container.winfo_children()
        for child in children:
            child.pack_forget()
        for child in children:
            if child is self.overflow_btn:
                continue
            if isinstance(child, ctk.CTkButton):
                if visible_count < MAX_VISIBLE:
                    child.pack(side="left", padx=2)
                    # Only count running (non-pinned) buttons toward MAX_VISIBLE
                    if not getattr(child, "_is_pinned", False):
                        visible_count += 1
                else:
                    self.overflow_apps.append(child)
            elif isinstance(child, ctk.CTkFrame):
                # Separator — always re-pack at consistent padx
                child.pack(side="left", padx=7)

        if self.overflow_apps:
            self.overflow_btn.pack(side="left", padx=2)
        elif hasattr(self, "overflow_menu") and self.overflow_menu.winfo_exists():
            self.overflow_menu.destroy()
        self.after(500, self.update_open_apps)

    def check_fullscreen(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                self.after(500, self.check_fullscreen)
                return

            # Cache my_hwnd — it never changes after the window is created
            if self._my_hwnd is None:
                self._my_hwnd = ctypes.windll.user32.GetAncestor(
                    self.winfo_id(), GA_ROOT
                )
            my_hwnd = self._my_hwnd

            is_me = hwnd == my_hwnd
            if not is_me:
                for attr in (
                    "search_window",
                    "tray_menu",
                    "overflow_menu",
                    "start_menu",
                ):
                    popup = getattr(self, attr, None)
                    if popup and popup.winfo_exists():
                        if hwnd == ctypes.windll.user32.GetAncestor(
                            popup.winfo_id(), GA_ROOT
                        ):
                            is_me = True
                            break

            HWND_BOTTOM = 1
            HWND_TOPMOST = -1
            FLAGS = 0x0002 | 0x0001 | 0x0010
            if is_me:
                ctypes.windll.user32.SetWindowPos(
                    my_hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
                )
                self.attributes("-alpha", float(self.config.get("opacity", 1.0)))
                self.after(500, self.check_fullscreen)
                return

            screen_w = ctypes.windll.user32.GetSystemMetrics(0)
            screen_h = ctypes.windll.user32.GetSystemMetrics(1)
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            is_fullscreen = (
                win32gui.GetClassName(hwnd) not in ("WorkerW", "Progman")
                and left <= 0
                and top <= 0
                and right >= screen_w
                and bottom >= screen_h
            )

            if is_fullscreen:
                ctypes.windll.user32.SetWindowPos(
                    my_hwnd, HWND_BOTTOM, 0, 0, 0, 0, FLAGS
                )
                self.attributes("-alpha", 0.0)
            else:
                ctypes.windll.user32.SetWindowPos(
                    my_hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
                )
                self.attributes("-alpha", float(self.config.get("opacity", 1.0)))
        except Exception:
            pass
        self.after(500, self.check_fullscreen)

    def focus_window(self, hwnd):
        if self.edit_mode:
            return
        try:
            if win32gui.GetForegroundWindow() == hwnd and not win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            else:
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def _close_window(self, hwnd):
        try:
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass

    def _fade_in(self, window, target_alpha: float):
        """Animate a toplevel from invisible to target_alpha over ~150ms."""
        window.wm_attributes("-alpha", 0.0)
        step = target_alpha / 10

        def _tick(current):
            if not window.winfo_exists():
                return
            nxt = min(current + step, target_alpha)
            window.wm_attributes("-alpha", nxt)
            if nxt < target_alpha:
                window.after(15, lambda: _tick(nxt))

        window.after(5, lambda: _tick(0.0))

    def _show_tooltip(self, widget, text):
        self._hide_tooltip()
        tip = tk.Toplevel(self)
        tip.overrideredirect(True)
        tip.wm_attributes("-topmost", True)
        tip.configure(bg="#2a2a2a")

        # Plain tk widgets — synchronous creation, no CTkToplevel delayed init
        border = tk.Frame(tip, bg="#3c3c3c", padx=1, pady=1)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg="#2a2a2a", padx=9, pady=5)
        inner.pack(fill="both", expand=True)
        tk.Label(
            inner, text=text, font=("Segoe UI Variable", 11), fg="#d8d8d8", bg="#2a2a2a"
        ).pack()

        tip.update_idletasks()

        # Hide from Alt+Tab
        hwnd = ctypes.windll.user32.GetAncestor(tip.winfo_id(), GA_ROOT)
        exstyle = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(
            hwnd, GWL_EXSTYLE, (exstyle & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
        )

        tw = tip.winfo_reqwidth()
        th = tip.winfo_reqheight()
        tx = widget.winfo_rootx() + widget.winfo_width() // 2 - tw // 2
        if self.config.get("position", "Top") == "Bottom":
            ty = widget.winfo_rooty() - th - 8
        else:
            ty = widget.winfo_rooty() + widget.winfo_height() + 8
        tip.geometry(f"+{tx}+{ty}")
        self._tooltip = tip

    def _schedule_tooltip(self, widget, text):
        self._cancel_tooltip()
        self._tooltip_after = self.after(500, lambda: self._show_tooltip(widget, text))

    def _cancel_tooltip(self):
        after_id = getattr(self, "_tooltip_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
            self._tooltip_after = None

    def _hide_tooltip(self):
        self._cancel_tooltip()
        tip = getattr(self, "_tooltip", None)
        if tip is not None:
            try:
                tip.destroy()
            except Exception:
                pass
            self._tooltip = None

    def toggle_custom_start(self):
        if self.edit_mode:
            return
        if hasattr(self, "start_menu") and self.start_menu.winfo_exists():
            self.start_menu.destroy()
            return

        self.start_menu = ctk.CTkToplevel(self)
        self.start_menu.overrideredirect(True)
        self.start_menu.attributes("-topmost", True)
        CK = "#000001"
        self.start_menu.configure(fg_color=CK)
        self.start_menu.wm_attributes(
            "-transparentcolor", CK, "-alpha", float(self.config.get("opacity", 1.0))
        )

        menu_w = 300
        menu_h = 400
        btn_x = self.start_btn.winfo_rootx()
        btn_y = self.start_btn.winfo_rooty()

        if self.config.get("position") == "Bottom":
            pos_y = btn_y - menu_h - 10
        else:
            pos_y = btn_y + self.start_btn.winfo_height() + 10

        self.start_menu.geometry(f"{menu_w}x{menu_h}+{btn_x}+{pos_y}")
        main_frame = ctk.CTkFrame(
            self.start_menu,
            corner_radius=15,
            border_width=1,
            border_color="#3c3c3c",
            fg_color="#1c1c1c",
        )
        main_frame.pack(fill="both", expand=True)

        current_user = os.environ.get("USERNAME", "User")
        profile_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        profile_frame.pack(fill="x", padx=15, pady=(15, 10))
        ctk.CTkLabel(
            profile_frame,
            text="\ue77b",
            font=("Segoe MDL2 Assets", 20),
            text_color="#9d9d9d",
        ).pack(side="left")
        ctk.CTkLabel(
            profile_frame,
            text=f"  {current_user}",
            font=("Segoe UI Variable", 18, "bold"),
        ).pack(side="left")

        self.start_scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        self.start_scroll.pack(fill="both", expand=True, padx=10, pady=5)

        pinned_apps = self.config.get("pinned_apps", [])
        if not pinned_apps:
            ctk.CTkLabel(
                self.start_scroll,
                text="\ue71d",
                font=("Segoe MDL2 Assets", 28),
                text_color="#444444",
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.start_scroll,
                text="No pinned apps",
                font=("Segoe UI Variable", 13, "bold"),
                text_color="#555555",
            ).pack()
            ctk.CTkLabel(
                self.start_scroll,
                text="Add entries to config.json",
                font=("Segoe UI Variable", 11),
                text_color="#444444",
            ).pack(pady=(2, 0))
        else:
            for app in pinned_apps:
                if os.path.exists(app["path"]):
                    img = get_icon_from_exe(app["path"])
                    icon = ctk.CTkImage(img, img, size=(32, 32)) if img else None
                    btn = ctk.CTkButton(
                        self.start_scroll,
                        text=f"  {app['name']}",
                        image=icon,
                        anchor="w",
                        height=48,
                        fg_color="transparent",
                        hover_color="#303030",
                        corner_radius=8,
                        font=("Segoe UI Variable Semibold", 14),
                        command=lambda p=app["path"]: [
                            subprocess.Popen([p]),
                            self.start_menu.destroy(),
                        ],
                    )
                    btn.pack(fill="x", pady=2, padx=5)

        bottom_bar = ctk.CTkFrame(main_frame, fg_color="#161616", corner_radius=10)
        bottom_bar.pack(fill="x", side="bottom", padx=10, pady=10)

        settings_btn = ctk.CTkButton(
            bottom_bar,
            text="⚙  Settings",
            width=115,
            fg_color="transparent",
            hover_color="#303030",
            anchor="w",
            font=("Segoe UI Variable", 13),
            command=lambda: [
                os.system("start ms-settings:"),
                self.start_menu.destroy(),
            ],
        )
        settings_btn.pack(side="left", padx=5, pady=5)
        shutdown_btn = ctk.CTkButton(
            bottom_bar,
            text="\ue7e8",
            width=32,
            height=32,
            fg_color="#c42b1c",
            hover_color="#a32012",
            corner_radius=8,
            font=("Segoe MDL2 Assets", 14),
            command=lambda: os.system("shutdown /s /t 0"),
        )
        shutdown_btn.pack(side="right", padx=5, pady=5)
        sleep_btn = ctk.CTkButton(
            bottom_bar,
            text="\uec46",
            width=32,
            height=32,
            fg_color="#2d2d2d",
            hover_color="#404040",
            corner_radius=8,
            font=("Segoe MDL2 Assets", 14),
            command=lambda: os.system(
                "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"
            ),
        )
        sleep_btn.pack(side="right", padx=5, pady=5)

        def on_focus_out(event):
            mx, my = self.winfo_pointerxy()
            bx = self.start_btn.winfo_rootx()
            by = self.start_btn.winfo_rooty()
            bw = self.start_btn.winfo_width()
            bh = self.start_btn.winfo_height()
            if bx <= mx <= bx + bw and by <= my <= by + bh:
                return
            self.start_menu.destroy()

        self.start_menu.bind("<FocusOut>", on_focus_out)
        self._fade_in(self.start_menu, float(self.config.get("opacity", 1.0)))
        self.start_menu.focus_set()

    def toggle_search(self):
        if self.edit_mode:
            return
        if hasattr(self, "search_window") and self.search_window.winfo_exists():
            self.search_window.destroy()
            return

        self.search_window = ctk.CTkToplevel(self)
        self.search_window.overrideredirect(True)
        self.search_window.attributes("-topmost", True)
        CK = "#000001"
        self.search_window.configure(fg_color=CK)
        self.search_window.wm_attributes(
            "-transparentcolor", CK, "-alpha", float(self.config.get("opacity", 1.0))
        )

        search_w, search_h = 500, 450
        bar_center_x = self.winfo_x() + (self.winfo_width() // 2)
        sx = bar_center_x - (search_w // 2)
        if self.config.get("position", "Top") == "Bottom":
            sy = self.winfo_y() - search_h - 15
        else:
            sy = self.winfo_y() + self.winfo_height() + 15

        self.search_window.geometry(f"{search_w}x{search_h}+{sx}+{sy}")
        main_frame = ctk.CTkFrame(
            self.search_window,
            corner_radius=22,
            border_width=1,
            border_color="#404040",
            fg_color="#1c1c1c",
        )
        main_frame.pack(fill="both", expand=True)
        search_area = ctk.CTkFrame(main_frame, fg_color="#252525", corner_radius=12)
        search_area.pack(fill="x", padx=15, pady=(15, 10))
        ctk.CTkLabel(
            search_area,
            text="\ue721",
            font=("Segoe MDL2 Assets", 14),
            text_color="#8a8a8a",
        ).pack(side="left", padx=(12, 0))
        self.search_entry = ctk.CTkEntry(
            search_area,
            placeholder_text="Search apps...",
            width=400,
            height=45,
            font=("Segoe UI Variable", 16),
            fg_color="transparent",
            border_width=0,
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.search_entry.focus_force()
        self.results_scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        self.results_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.update_search_results)
        self.search_window.bind("<Escape>", lambda e: self.search_window.destroy())
        self.search_entry.bind("<Return>", lambda e: self.launch_top_result())

        def on_search_focus_out(event):
            mx, my = self.winfo_pointerxy()
            bx = self.search_btn.winfo_rootx()
            by = self.search_btn.winfo_rooty()
            bw = self.search_btn.winfo_width()
            bh = self.search_btn.winfo_height()
            if bx <= mx <= bx + bw and by <= my <= by + bh:
                return
            self.search_window.after(100, _check_search_focus)

        def _check_search_focus():
            try:
                if not self.search_window.winfo_exists():
                    return
                focused = self.focus_get()
                if focused is None:
                    self.search_window.destroy()
                    return
                if not str(focused).startswith(str(self.search_window)):
                    self.search_window.destroy()
            except Exception:
                pass

        self._fade_in(self.search_window, float(self.config.get("opacity", 1.0)))
        self.search_window.bind("<FocusOut>", on_search_focus_out)

    def update_search_results(self, event=None):
        query = self.search_entry.get()
        for child in self.results_scroll.winfo_children():
            child.destroy()
        if len(query) < 1:
            return
        apps = search_windows_apps(query)
        if not apps:
            ctk.CTkLabel(
                self.results_scroll,
                text="\ue721",
                font=("Segoe MDL2 Assets", 28),
                text_color="#3a3a3a",
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.results_scroll,
                text="No results found",
                font=("Segoe UI Variable", 13),
                text_color="#666666",
            ).pack()
            return
        for app in apps:
            exe = (
                resolve_lnk(app["path"])
                if app["path"].endswith(".lnk")
                else app["path"]
            )
            img = get_icon_from_exe(exe) if exe else None
            icon = ctk.CTkImage(img, img, size=(28, 28)) if img else None
            ctk.CTkButton(
                self.results_scroll,
                text=f"  {app['name']}",
                image=icon,
                anchor="w",
                height=46,
                border_width=0,
                corner_radius=8,
                fg_color="transparent",
                hover_color="#2a2a2a",
                font=("Segoe UI Variable Semibold", 13),
                command=lambda p=app["path"]: [
                    os.startfile(p),
                    self.search_window.destroy(),
                ],
            ).pack(fill="x", pady=2, padx=5)

    def launch_top_result(self):
        query = self.search_entry.get()
        if len(query) < 1:
            return
        apps = search_windows_apps(query)
        if apps:
            os.startfile(apps[0]["path"])
            if hasattr(self, "search_window") and self.search_window.winfo_exists():
                self.search_window.destroy()

    def toggle_control_center(self):
        if self.edit_mode:
            return
        if hasattr(self, "tray_menu") and self.tray_menu.winfo_exists():
            self.tray_menu.destroy()
            return

        self.tray_menu = ctk.CTkToplevel(self)
        self.tray_menu.overrideredirect(True)
        CK = "#000001"
        self.tray_menu.configure(fg_color=CK)
        self.tray_menu.wm_attributes(
            "-transparentcolor",
            CK,
            "-alpha",
            float(self.config.get("opacity", 1.0)),
            "-topmost",
            True,
        )
        f = ctk.CTkFrame(
            self.tray_menu,
            corner_radius=20,
            fg_color="#1c1c1c",
            border_width=1,
            border_color="#3c3c3c",
        )
        f.pack(fill="both", expand=True)

        menu_w, menu_h = 250, 300
        btn_x = self.tray_btn.winfo_rootx()
        btn_y = self.tray_btn.winfo_rooty()
        btn_w = self.tray_btn.winfo_width()
        btn_h = self.tray_btn.winfo_height()

        pos_x = btn_x + btn_w - menu_w
        if self.config.get("position", "Top") == "Bottom":
            pos_y = btn_y - menu_h - 10
        else:
            pos_y = btn_y + btn_h + 10

        self.tray_menu.geometry(f"{menu_w}x{menu_h}+{pos_x}+{pos_y}")

        # Header
        header = ctk.CTkFrame(f, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 4))
        ctk.CTkLabel(
            header, text="\ue782", font=("Segoe MDL2 Assets", 13), text_color="#9d9d9d"
        ).pack(side="left")
        ctk.CTkLabel(
            header, text="  WinBar", font=("Segoe UI Variable", 15, "bold")
        ).pack(side="left")

        # Divider
        ctk.CTkFrame(f, height=1, fg_color="#2e2e2e").pack(
            fill="x", padx=12, pady=(0, 8)
        )

        # Opacity
        ctk.CTkLabel(
            f, text="Opacity", font=("Segoe UI Variable", 11), text_color="#8a8a8a"
        ).pack(anchor="w", padx=18)
        ctk.CTkSlider(
            f,
            from_=0.1,
            to=1.0,
            number_of_steps=18,
            command=self.change_opacity,
        ).pack(pady=(2, 10), padx=18, fill="x")

        # Position
        ctk.CTkLabel(
            f, text="Position", font=("Segoe UI Variable", 11), text_color="#8a8a8a"
        ).pack(anchor="w", padx=18)
        pos_menu = ctk.CTkOptionMenu(
            f, values=["Top", "Bottom"], command=self.change_position, width=214
        )
        pos_menu.set(self.config.get("position", "Top"))
        pos_menu.pack(pady=(2, 10), padx=18)

        # Actions
        ctk.CTkButton(
            f,
            text="Edit Layout",
            fg_color="#0078D4",
            hover_color="#106EBE",
            height=32,
            command=self.toggle_edit_mode,
        ).pack(padx=18, pady=(0, 6), fill="x")
        ctk.CTkButton(
            f,
            text="Quit WinBar",
            fg_color="#c42b1c",
            hover_color="#a32012",
            height=32,
            command=self.safe_exit,
        ).pack(padx=18, pady=(0, 14), fill="x")

        def on_tray_focus_out(event):
            mx, my = self.winfo_pointerxy()
            bx = self.tray_btn.winfo_rootx()
            by = self.tray_btn.winfo_rooty()
            bw = self.tray_btn.winfo_width()
            bh = self.tray_btn.winfo_height()
            if bx <= mx <= bx + bw and by <= my <= by + bh:
                return
            self.tray_menu.destroy()

        self._fade_in(self.tray_menu, float(self.config.get("opacity", 1.0)))
        self.tray_menu.bind("<FocusOut>", on_tray_focus_out)
        self.tray_menu.focus_set()

    def change_opacity(self, v):
        self.wm_attributes("-alpha", v)
        self.config["opacity"] = v
        self.save_layout()

    def change_position(self, value):
        self.config["position"] = value
        self.save_layout()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * (float(self.config.get("width_percent", 90)) / 100))
        h = int(self.config.get("height", 50))
        x_pos = (sw - w) // 2
        if value == "Bottom":
            y_pos = sh - h - 10
            edge = ABE_BOTTOM
        else:
            y_pos = 10
            edge = ABE_TOP

        self.geo_string = f"{w}x{h}+{x_pos}+{y_pos}"
        self.geometry(self.geo_string)
        self.update()

        HWND_TOPMOST = -1
        my_hwnd = ctypes.windll.user32.GetAncestor(self.winfo_id(), GA_ROOT)
        ctypes.windll.user32.SetWindowPos(
            my_hwnd, HWND_TOPMOST, x_pos, y_pos, w, h, 0x0040
        )
        unregister_appbar()
        register_appbar(self.winfo_id(), h, -5, edge)

        if hasattr(self, "tray_menu") and self.tray_menu.winfo_exists():
            self.tray_menu.destroy()
            self.toggle_control_center()

    def toggle_overflow_menu(self):
        if self.edit_mode:
            return
        if hasattr(self, "overflow_menu") and self.overflow_menu.winfo_exists():
            self.overflow_menu.destroy()
            return
        if not self.overflow_apps:
            return

        self.overflow_menu = ctk.CTkToplevel(self)
        self.overflow_menu.overrideredirect(True)
        self.overflow_menu.attributes("-topmost", True)
        CK = "#000001"
        self.overflow_menu.configure(fg_color=CK)
        self.overflow_menu.wm_attributes(
            "-transparentcolor", CK, "-alpha", float(self.config.get("opacity", 1.0))
        )

        menu_h = min(len(self.overflow_apps) * 50 + 20, 500)
        menu_w = 220
        btn_x = self.overflow_btn.winfo_rootx()
        btn_y = self.overflow_btn.winfo_rooty()

        if self.config.get("position", "Top") == "Bottom":
            pos_y = btn_y - menu_h - 10
        else:
            pos_y = btn_y + self.overflow_btn.winfo_height() + 10

        self.overflow_menu.geometry(f"{menu_w}x{menu_h}+{btn_x}+{pos_y}")
        main_frame = ctk.CTkFrame(
            self.overflow_menu,
            corner_radius=15,
            border_width=1,
            border_color="#3c3c3c",
            fg_color="#1c1c1c",
        )
        main_frame.pack(fill="both", expand=True)

        scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        scroll.pack(fill="both", expand=True, padx=5, pady=5)

        for original_btn in self.overflow_apps:
            icon = original_btn.cget("image")
            cmd = original_btn.cget("command")
            full_name = getattr(original_btn, "_full_name", "App")
            btn = ctk.CTkButton(
                scroll,
                text=f"  {full_name[:25]}",
                image=icon,
                anchor="w",
                height=40,
                fg_color="transparent",
                hover_color="#303030",
                font=("Segoe UI Variable Semibold", 13),
                command=lambda c=cmd: (
                    [c(), self.overflow_menu.destroy()] if c else None
                ),
            )
            btn.pack(fill="x", pady=2)

        self._fade_in(self.overflow_menu, float(self.config.get("opacity", 1.0)))
        self.overflow_menu.bind("<FocusOut>", lambda e: self.overflow_menu.destroy())
        self.overflow_menu.focus_set()

    def _set_cursor_recursive(self, widget, cursor):
        try:
            widget.configure(cursor=cursor)
        except Exception:
            pass
        for child in widget.winfo_children():
            self._set_cursor_recursive(child, cursor)

    def toggle_edit_mode(self, event=None):
        self.edit_mode = not self.edit_mode

        # Close any open popups before entering/leaving edit mode
        for attr in ("start_menu", "search_window", "tray_menu", "overflow_menu"):
            popup = getattr(self, attr, None)
            if popup and popup.winfo_exists():
                popup.destroy()

        if self.edit_mode:
            self.pill_frame.configure(
                fg_color=self.config.get("bg_color"),
                border_color="#f0a500",
                border_width=2,
            )
            self.active_window_label.configure(
                text="\ue70f  Edit Mode  —  drag to rearrange",
                text_color="#f0a500",
            )
            self._set_cursor_recursive(self.pill_frame, "fleur")
            # Show draggable zone borders on all widgets
            for w in self.widget_map.values():
                try:
                    w.configure(border_width=1, border_color="#5a4010")
                except Exception:
                    pass
            self.edit_done_btn.pack(in_=self.right_wing, side="right", padx=(4, 2))
        else:
            self._cancel_drag()  # clean up any in-progress drag
            self.pill_frame.configure(
                fg_color=self.config.get("bg_color"),
                border_color="#3a3a3a",
                border_width=1,
            )
            self.active_window_label.configure(text="", text_color="#b0b0b0")
            self._set_cursor_recursive(self.pill_frame, "arrow")
            # Clear all widget borders
            for w in self.widget_map.values():
                try:
                    w.configure(border_width=0)
                except Exception:
                    pass
            self.edit_done_btn.pack_forget()
            # Restore active indicator if needed
            if self._active_hwnd in self.active_app_buttons:
                btn = self.active_app_buttons[self._active_hwnd]
                if btn.winfo_exists():
                    btn.configure(
                        fg_color="#1a3a52", border_width=1, border_color="#0078D4"
                    )
            self.save_layout()

    def save_layout(self):
        with open(self.config_path, "w") as f:
            json.dump(self.config, f, indent=4)

    # ── Drag-and-drop helpers ─────────────────────────────────────────────────

    _WIDGET_LABELS = {
        "start": "Start",
        "search": "Search",
        "apps": "Apps",
        "active_window": "Active Window",
        "taskmanager": "Task Manager",
        "sys_monitor": "System Monitor",
        "clock": "Clock",
        "tray": "Tray",
    }

    def _widget_name_at(self, x, y):
        """Precise bounding-box hit-test against widget_map items.
        Avoids winfo_containing which resolves to the deepest child widget."""
        for key, w in self.widget_map.items():
            try:
                wx, wy = w.winfo_rootx(), w.winfo_rooty()
                ww, wh = w.winfo_width(), w.winfo_height()
                if wx <= x <= wx + ww and wy <= y <= wy + wh:
                    return key
            except Exception:
                pass
        return None

    def _create_drag_ghost(self, name):
        """Create a small floating chip that follows the cursor while dragging."""
        label = self._WIDGET_LABELS.get(name, name.replace("_", " ").title())
        ghost = tk.Toplevel(self)
        ghost.overrideredirect(True)
        ghost.attributes("-topmost", True)
        ghost.attributes("-alpha", 0.88)
        ghost.configure(bg="#f0a500")
        outer = tk.Frame(ghost, bg="#f0a500", padx=10, pady=5)
        outer.pack()
        tk.Label(
            outer,
            text=label,
            font=("Segoe UI Variable", 10, "bold"),
            fg="#1a1000",
            bg="#f0a500",
        ).pack()
        ghost.update_idletasks()
        # Hide from Alt+Tab
        hwnd = ctypes.windll.user32.GetAncestor(ghost.winfo_id(), GA_ROOT)
        ex = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        ctypes.windll.user32.SetWindowLongW(
            hwnd, GWL_EXSTYLE, (ex & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
        )
        return ghost

    def _cancel_drag(self):
        """Tear down any in-progress drag without applying it."""
        ghost = self.drag_data.get("ghost")
        if ghost:
            try:
                ghost.destroy()
            except Exception:
                pass
        self.drag_data.update(
            {"widget_name": None, "hover_target": None, "ghost": None}
        )

    def _restore_zone_borders(self):
        """Reset every widget_map item to the dim 'draggable zone' border."""
        for w in self.widget_map.values():
            try:
                w.configure(border_width=1, border_color="#5a4010")
            except Exception:
                pass

    # ── Drag event handlers ───────────────────────────────────────────────────

    def _on_drag_start(self, event, name):
        if not self.edit_mode:
            return
        self._hide_tooltip()
        self.drag_data["widget_name"] = name
        self.drag_data["hover_target"] = None

        # Highlight the source widget in amber
        src = self.widget_map.get(name)
        if src:
            try:
                src.configure(border_width=2, border_color="#f0a500")
            except Exception:
                pass

        # Create ghost chip and place it just to the side of the cursor
        ghost = self._create_drag_ghost(name)
        x, y = self.winfo_pointerxy()
        ghost.geometry(f"+{x + 14}+{y - 22}")
        self.drag_data["ghost"] = ghost

    def _on_drag_motion(self, event, name):
        if not self.edit_mode or not self.drag_data.get("widget_name"):
            return
        source = self.drag_data["widget_name"]
        x, y = self.winfo_pointerxy()

        # Keep ghost chip glued to cursor
        ghost = self.drag_data.get("ghost")
        if ghost and ghost.winfo_exists():
            ghost.geometry(f"+{x + 14}+{y - 22}")

        # Bbox hit-test for drop target
        target = self._widget_name_at(x, y)
        prev = self.drag_data.get("hover_target")

        if target == prev:
            return  # nothing changed

        # Restore old hover target to dim border
        if prev and prev != source:
            pw = self.widget_map.get(prev)
            if pw:
                try:
                    pw.configure(border_width=1, border_color="#5a4010")
                except Exception:
                    pass

        # Highlight new drop target in light blue
        if target and target != source:
            tw = self.widget_map.get(target)
            if tw:
                try:
                    tw.configure(border_width=2, border_color="#4fc3f7")
                except Exception:
                    pass

        self.drag_data["hover_target"] = (
            target if (target and target != source) else None
        )

    def _on_drag_drop(self, event, name):
        if not self.edit_mode or not self.drag_data.get("widget_name"):
            return
        source = self.drag_data["widget_name"]
        x, y = self.winfo_pointerxy()
        target = self._widget_name_at(x, y)

        # Tear down ghost immediately
        ghost = self.drag_data.get("ghost")
        if ghost:
            try:
                ghost.destroy()
            except Exception:
                pass
        self.drag_data.update(
            {"widget_name": None, "hover_target": None, "ghost": None}
        )

        # Restore all zone borders first
        self._restore_zone_borders()

        if not target or source == target:
            return

        layout = self.config.get("layout", {})
        source_zone = target_zone = None
        source_idx = target_idx = -1

        for zone_name, items in layout.items():
            if source in items:
                source_zone, source_idx = zone_name, items.index(source)
            if target in items:
                target_zone, target_idx = zone_name, items.index(target)

        if source_zone and target_zone and source_idx != -1 and target_idx != -1:
            layout[source_zone].pop(source_idx)
            layout[target_zone].insert(target_idx, source)
            self.config["layout"] = layout
            self.render_bar()
            self._set_cursor_recursive(self.pill_frame, "fleur")
            self._restore_zone_borders()

    def safe_exit(self):
        set_taskbar_visibility(True)
        unregister_appbar()
        self.quit()


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()

    set_taskbar_visibility(False)
    reset_work_area()
    atexit.register(unregister_appbar)
    atexit.register(lambda: set_taskbar_visibility(True))

    app = FloatingBar()
    app.protocol("WM_DELETE_WINDOW", app.safe_exit)
    app.update_idletasks()

    root = ctypes.windll.user32.GetAncestor(app.winfo_id(), GA_ROOT)
    style = ctypes.windll.user32.GetWindowLongW(root, GWL_EXSTYLE)
    ctypes.windll.user32.SetWindowLongW(
        root, GWL_EXSTYLE, (style & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
    )

    edge = ABE_BOTTOM if app.config.get("position", "Top") == "Bottom" else ABE_TOP
    register_appbar(app.winfo_id(), 50, -5, edge)

    app.geometry(app.geo_string)
    app.update_time()
    app._poll_foreground()
    app.update_open_apps()
    threading.Thread(target=install_to_startup, daemon=True).start()
    threading.Thread(target=setup_system_tray, args=(app,), daemon=True).start()
    app.mainloop()
