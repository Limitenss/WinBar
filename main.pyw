import os
import json
import datetime
import customtkinter as ctk
import win32gui
import win32con
import psutil
import pystray
from PIL import Image, ImageDraw
import threading
import winreg
import sys
import subprocess
import ctypes
from ctypes import wintypes
import atexit
import win32api
import win32ui
import win32process
import requests

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


def search_windows_apps(query):
    search_paths = [
        os.path.join(
            os.environ["ProgramData"], "Microsoft", "Windows", "Start Menu", "Programs"
        ),
        os.path.join(
            os.environ["AppData"], "Microsoft", "Windows", "Start Menu", "Programs"
        ),
    ]
    results = []
    query = query.lower()
    for path in search_paths:
        if os.path.exists(path):
            for root, dirs, files in os.walk(path):
                for file in files:
                    if file.endswith(".lnk") and query in file.lower():
                        results.append(
                            {"name": file[:-4], "path": os.path.join(root, file)}
                        )
    return results[:7]


def set_taskbar_visibility(visible=True):
    hwnd_tray = win32gui.FindWindow("Shell_TrayWnd", None)

    class APPBARDATA_STATE(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("hWnd", wintypes.HWND),
            ("uCallbackMessage", wintypes.UINT),
            ("uEdge", wintypes.UINT),
            ("rc", wintypes.RECT),
            ("lParam", wintypes.LPARAM),
        ]

    abd = APPBARDATA_STATE()
    abd.cbSize = ctypes.sizeof(abd)
    abd.hWnd = hwnd_tray
    abd.lParam = 1 if not visible else 2
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


def get_exe_from_hwnd(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return psutil.Process(pid).exe()
    except Exception:
        return None


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


def get_new_data():
    now = datetime.datetime.now()
    return f"{now.strftime('%#I:%M %p')}  •  {now.strftime('%a, %b %#d')}"


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


def get_running_apps():
    hwnds = []
    win32gui.EnumWindows(lambda hwnd, param: param.append(hwnd), hwnds)
    valid = []
    blacklist = [
        "Program Manager",
        "Microsoft Text Input Application",
        "NVIDIA GeForce Overlay",
        "Discord Updater",
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
            border_color="lightblue",
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
                "left": ["apps", "active_window"],
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

        self.active_window_label = ctk.CTkLabel(
            self.pill_frame,
            text="",
            font=("Segoe UI Variable", 12, "bold"),
            text_color="#3d5a80",
        )
        self.my_label = ctk.CTkLabel(
            self.pill_frame,
            text=get_new_data(),
            font=("Segoe UI Variable", 20),
            text_color="white",
        )

        self.sys_container = ctk.CTkFrame(
            self.pill_frame, fg_color="transparent", width=180, height=40
        )
        self.sys_container.pack_propagate(False)
        self.sys_label = ctk.CTkLabel(
            self.sys_container,
            text="CPU: 0% | RAM: 0%",
            font=("Segoe UI Variable", 12),
            text_color="#A0A0A0",
        )
        self.sys_label.pack(expand=True, fill="both")

        self.apps_container = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        btn_size = int(h * 0.75)
        self.overflow_btn = ctk.CTkButton(
            self.apps_container,
            text="⋮",
            width=30,
            height=btn_size,
            fg_color="#1a3b5c",
            hover_color="#3d5a80",
            corner_radius=10,
            font=("Segoe UI Variable", 20, "bold"),
            command=self.toggle_overflow_menu,
        )
        self.overflow_apps = []

        self.tray_btn = ctk.CTkButton(
            self.pill_frame,
            text="⚙️",
            width=30,
            fg_color="transparent",
            hover_color="#2a2a2a",
            font=("Segoe UI Emoji", 16),
            command=self.toggle_control_center,
        )
        self.search_btn = ctk.CTkButton(
            self.pill_frame,
            text="🔍",
            width=30,
            fg_color="transparent",
            hover_color="#2a2a2a",
            font=("Segoe UI Emoji", 16),
            command=self.toggle_search,
        )
        self.taskmanager_btn = ctk.CTkButton(
            self.pill_frame,
            text="📊",
            width=30,
            fg_color="transparent",
            hover_color="#2a2a2a",
            font=("Segoe UI Emoji", 16),
            command=lambda: subprocess.Popen("start taskmgr", shell=True),
        )

        self.widget_map = {
            "search": self.search_btn,
            "apps": self.apps_container,
            "active_window": self.active_window_label,
            "taskmanager": self.taskmanager_btn,
            "sys_monitor": self.sys_container,
            "clock": self.my_label,
            "tray": self.tray_btn,
        }

        self.edit_mode = False
        self.drag_data = {"widget_name": None}
        self.bind_all("<Control-space>", lambda e: self.toggle_search())
        self.bind("<Control-e>", self.toggle_edit_mode)

        for name, widget in self.widget_map.items():
            widget.bind(
                "<ButtonPress-1>", lambda e, n=name: self.on_drag_start(e, n), add="+"
            )
            widget.bind(
                "<ButtonRelease-1>", lambda e, n=name: self.on_drag_drop(e, n), add="+"
            )

        psutil.cpu_percent(interval=None)
        self.render_bar()
        self.check_fullscreen()

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
            border_color="lightblue",
            fg_color="#1a1a1b",
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
                hover_color="#2a2a2a",
                font=("Segoe UI Variable Semibold", 13),
                command=lambda c=cmd: (
                    [c(), self.overflow_menu.destroy()] if c else None
                ),
            )
            btn.pack(fill="x", pady=2)

        self.overflow_menu.bind("<FocusOut>", lambda e: self.overflow_menu.destroy())
        self.overflow_menu.focus_set()

    def check_fullscreen(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                self.after(500, self.check_fullscreen)
                return

            my_hwnd = ctypes.windll.user32.GetAncestor(self.winfo_id(), GA_ROOT)
            is_me = hwnd == my_hwnd

            if hasattr(self, "search_window") and self.search_window.winfo_exists():
                if hwnd == ctypes.windll.user32.GetAncestor(
                    self.search_window.winfo_id(), GA_ROOT
                ):
                    is_me = True

            if hasattr(self, "tray_menu") and self.tray_menu.winfo_exists():
                if hwnd == ctypes.windll.user32.GetAncestor(
                    self.tray_menu.winfo_id(), GA_ROOT
                ):
                    is_me = True

            if hasattr(self, "overflow_menu") and self.overflow_menu.winfo_exists():
                if hwnd == ctypes.windll.user32.GetAncestor(
                    self.overflow_menu.winfo_id(), GA_ROOT
                ):
                    is_me = True

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
            rect = win32gui.GetWindowRect(hwnd)
            left, top, right, bottom = rect

            class_name = win32gui.GetClassName(hwnd)
            if class_name in ["WorkerW", "Progman"]:
                is_fullscreen = False
            else:
                is_fullscreen = (
                    left <= 0 and top <= 0 and right >= screen_w and bottom >= screen_h
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

    def toggle_search(self):
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
            border_width=2,
            border_color="#3d5a80",
            fg_color="#141415",
        )
        main_frame.pack(fill="both", expand=True)
        search_area = ctk.CTkFrame(main_frame, fg_color="#1c1c1e", corner_radius=15)
        search_area.pack(fill="x", padx=15, pady=(15, 10))
        ctk.CTkLabel(
            search_area, text=" 🔍 ", font=("Segoe UI Emoji", 16), text_color="#3d5a80"
        ).pack(side="left", padx=(10, 0))
        self.search_entry = ctk.CTkEntry(
            search_area,
            placeholder_text="Search apps...",
            width=400,
            height=45,
            font=("Segoe UI Variable", 18),
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
                text="No results found.",
                font=("Segoe UI Variable", 13),
                text_color="#777777",
            ).pack(pady=20)
            return
        for app in apps:
            img = get_icon_from_exe(app["path"])
            icon = ctk.CTkImage(img, img, size=(28, 28)) if img else None
            ctk.CTkButton(
                self.results_scroll,
                text=f"  {app['name']}",
                image=icon,
                anchor="w",
                height=55,
                border_width=1,
                border_color="lightblue",
                corner_radius=12,
                fg_color="transparent",
                hover_color="#252526",
                font=("Segoe UI Variable Semibold", 15),
                command=lambda p=app["path"]: [
                    os.startfile(p),
                    self.search_window.destroy(),
                ],
            ).pack(fill="x", pady=3, padx=5)

    def launch_top_result(self):
        query = self.search_entry.get()
        if len(query) < 1:
            return
        apps = search_windows_apps(query)
        if apps:
            os.startfile(apps[0]["path"])
            if hasattr(self, "search_window") and self.search_window.winfo_exists():
                self.search_window.destroy()

    def focus_window(self, hwnd):
        if self.edit_mode:
            return
        try:
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def render_bar(self):
        for zone in [self.left_wing, self.center_wing, self.right_wing]:
            for child in zone.winfo_children():
                child.pack_forget()

        self.pill_frame.grid_columnconfigure(0, weight=1, uniform="side")
        self.pill_frame.grid_columnconfigure(1, weight=0)
        self.pill_frame.grid_columnconfigure(2, weight=1, uniform="side")

        self.left_wing.grid(row=0, column=0, sticky="w", padx=10)
        self.center_wing.grid(row=0, column=1)
        self.right_wing.grid(row=0, column=2, sticky="e", padx=10)

        layout = self.config.get(
            "layout",
            {
                "left": ["apps", "active_window"],
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
        for key in layout.get("right", []):
            self.add_widget_to_zone(key, self.right_wing)

    def add_widget_to_zone(self, key, zone):
        if key in self.widget_map:
            self.widget_map[key].pack(in_=zone, side="left", padx=5)

    def toggle_edit_mode(self, event=None):
        self.edit_mode = not self.edit_mode
        c = "#3b3014" if self.edit_mode else self.config.get("bg_color")
        b = "#ffae00" if self.edit_mode else "lightblue"
        self.pill_frame.configure(fg_color=c, border_color=b)
        if not self.edit_mode:
            self.save_layout()

    def on_drag_start(self, event, name):
        if self.edit_mode:
            self.drag_data["widget_name"] = name

    def on_drag_drop(self, event, name):
        if not self.edit_mode or not self.drag_data["widget_name"]:
            return
        self.drag_data["widget_name"] = None

    def save_layout(self):
        with open(self.config_path, "w") as f:
            json.dump(self.config, f, indent=4)

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
            fg_color="#1a1a1b",
            border_width=1,
            border_color="lightblue",
        )
        f.pack(fill="both", expand=True)

        if self.config.get("position", "Top") == "Bottom":
            tray_y = self.winfo_y() - 260
        else:
            tray_y = self.winfo_y() + 60

        self.tray_menu.geometry(
            f"250x250+{(self.winfo_x() + self.winfo_width()) - 270}+{tray_y}"
        )
        ctk.CTkLabel(f, text="Settings", font=("Segoe UI Variable", 16, "bold")).pack(
            pady=10
        )
        ctk.CTkSlider(f, from_=0.1, to=1.0, command=self.change_opacity).pack(
            pady=10, padx=20
        )

        pos_menu = ctk.CTkOptionMenu(
            f,
            values=["Top", "Bottom"],
            command=self.change_position,
        )
        pos_menu.set(self.config.get("position", "Top"))
        pos_menu.pack(pady=5)

        ctk.CTkButton(
            f, text="Edit Layout", fg_color="#bd8100", command=self.toggle_edit_mode
        ).pack(pady=5)

        ctk.CTkButton(
            f, text="Quit App", fg_color="#d32f2f", command=self.safe_exit
        ).pack(pady=5)

    def change_opacity(self, v):
        self.wm_attributes("-alpha", v)
        self.config["opacity"] = v
        self.save_layout()

    def safe_exit(self):
        set_taskbar_visibility(True)
        unregister_appbar()
        self.quit()

    def update_time(self):
        self.my_label.configure(text=get_new_data())
        self.sys_label.configure(
            text=f"CPU: {int(psutil.cpu_percent()):2}% | RAM: {int(psutil.virtual_memory().percent)}%"
        )
        try:
            hwnd = win32gui.GetForegroundWindow()
            title = win32gui.GetWindowText(hwnd)
            if title:
                self.active_window_label.configure(text=f" > {title[:30]}...")
            else:
                self.active_window_label.configure(text="")
        except:
            pass
        self.after(1000, self.update_time)

    def update_open_apps(self):
        if not hasattr(self, "active_app_buttons"):
            self.active_app_buttons = {}
            pinned = self.config.get("pinned_apps", [])
            for app in pinned:
                if os.path.exists(app["path"]):
                    img = get_icon_from_exe(app["path"])
                    bar_h = int(self.config.get("height", 50))
                    icon = (
                        ctk.CTkImage(
                            img, img, size=(int(bar_h * 0.45), int(bar_h * 0.45))
                        )
                        if img
                        else None
                    )
                    p_btn = ctk.CTkButton(
                        self.apps_container,
                        text="" if icon else app["name"][:3],
                        image=icon,
                        width=int(bar_h * 0.75),
                        height=int(bar_h * 0.75),
                        fg_color="transparent",
                        command=lambda p=app["path"]: subprocess.Popen([p]),
                    )
                    p_btn._full_name = app["name"]
                    p_btn.pack(side="left", padx=2)

            if pinned:
                ctk.CTkFrame(
                    self.apps_container, width=2, height=20, fg_color="#555555"
                ).pack(side="left", padx=5)

        bar_h = int(self.config.get("height", 50))
        current_windows = get_running_apps()
        cur_hwnds = [app["hwnd"] for app in current_windows]

        for app in current_windows:
            if any(
                x in app["title"]
                for x in ["FloatingBar", "Limitens", "NVIDIA", "Overlay"]
            ):
                continue
            hid = app["hwnd"]
            if hid not in self.active_app_buttons:
                path = get_exe_from_hwnd(hid)
                img = get_icon_from_exe(path) if path else None
                icon = (
                    ctk.CTkImage(img, img, size=(int(bar_h * 0.45), int(bar_h * 0.45)))
                    if img
                    else None
                )

                btn = ctk.CTkButton(
                    self.apps_container,
                    text="" if icon else app["title"][:8],
                    image=icon,
                    width=int(bar_h * 0.75),
                    height=int(bar_h * 0.75),
                    fg_color="#1a3b5c",
                    hover_color="#3d5a80",
                    corner_radius=10,
                    command=lambda h=hid: self.focus_window(h),
                )
                btn._full_name = app["title"]
                btn.pack(side="left", padx=2)
                self.active_app_buttons[hid] = btn

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
            if hasattr(self, "overflow_btn") and child == self.overflow_btn:
                continue

            if isinstance(child, ctk.CTkButton):
                if visible_count < MAX_VISIBLE:
                    child.pack(side="left", padx=2)
                    visible_count += 1
                else:
                    self.overflow_apps.append(child)
            elif isinstance(child, ctk.CTkFrame):
                if visible_count < MAX_VISIBLE:
                    child.pack(side="left", padx=5)

        if self.overflow_apps:
            if hasattr(self, "overflow_btn"):
                self.overflow_btn.pack(side="left", padx=2)
        else:
            if hasattr(self, "overflow_menu") and self.overflow_menu.winfo_exists():
                self.overflow_menu.destroy()

        self.after(500, self.update_open_apps)


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
    app.update_open_apps()
    install_to_startup()
    threading.Thread(target=setup_system_tray, args=(app,), daemon=True).start()
    app.mainloop()
