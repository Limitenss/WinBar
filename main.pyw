import os, json, datetime, sys, subprocess, time
import tkinter as tk
from tkinter import filedialog
import customtkinter as ctk
import win32gui, win32con, win32ui, win32process
import psutil, pystray, threading, winreg, ctypes, atexit
from PIL import Image, ImageDraw
from ctypes import POINTER, byref, cast, c_float, wintypes

try:
    from comtypes import CLSCTX_ALL, COMMETHOD, GUID, HRESULT, IUnknown
    from comtypes import CoCreateInstance, CoInitialize, CoUninitialize

    _AUDIO_API_AVAILABLE = True
except Exception:
    CLSCTX_ALL = COMMETHOD = GUID = HRESULT = IUnknown = None
    CoCreateInstance = CoInitialize = CoUninitialize = None
    _AUDIO_API_AVAILABLE = False


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
WINDOW_TITLE_BLACKLIST = [
    "Program Manager",
    "Microsoft Text Input Application",
    "NVIDIA GeForce Overlay",
    "Discord Updater",
    "CTkToplevel",
]


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
    for h in hwnds:
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if not title or any(name in title for name in WINDOW_TITLE_BLACKLIST):
                continue
            style = win32gui.GetWindowLong(h, win32con.GWL_EXSTYLE)
            is_tool = style & win32con.WS_EX_TOOLWINDOW
            is_app = style & win32con.WS_EX_APPWINDOW
            if is_app and not is_tool:
                valid.append({"title": title, "hwnd": h})
            elif not is_tool and win32gui.GetParent(h) == 0:
                valid.append({"title": title, "hwnd": h})
    return valid


def get_background_apps(limit=14):
    visible_pids = set()
    blocked_pids = set()

    def _collect_visible(hwnd, param):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return
            title = win32gui.GetWindowText(hwnd)
            if not title:
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if any(name in title for name in WINDOW_TITLE_BLACKLIST):
                if pid:
                    blocked_pids.add(pid)
                return
            style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            if style & win32con.WS_EX_TOOLWINDOW:
                return
            if pid:
                visible_pids.add(pid)
        except Exception:
            pass

    try:
        win32gui.EnumWindows(_collect_visible, None)
    except Exception:
        pass

    skip_names = {
        "system idle process",
        "system",
        "registry",
        "memory compression",
        "svchost.exe",
        "fontdrvhost.exe",
        "dllhost.exe",
        "sihost.exe",
        "ctfmon.exe",
        "lockapp.exe",
        "searchhost.exe",
        "widgets.exe",
        "floatingbar.exe",
        "pythonw.exe",
    }
    entries = {}
    current_pid = os.getpid()
    current_user = os.environ.get("USERNAME", "").lower()

    for proc in psutil.process_iter(["pid", "name", "exe", "username", "memory_info"]):
        try:
            pid = proc.info["pid"]
            if pid == current_pid or pid in visible_pids or pid in blocked_pids:
                continue
            exe = proc.info.get("exe")
            if not exe or not os.path.exists(exe):
                continue
            name = (proc.info.get("name") or os.path.basename(exe)).lower()
            if name in skip_names:
                continue
            username = (proc.info.get("username") or "").lower()
            if current_user and current_user not in username:
                continue
            key = os.path.normcase(exe)
            item = entries.setdefault(
                key,
                {
                    "name": os.path.splitext(os.path.basename(exe))[0],
                    "path": exe,
                    "count": 0,
                    "memory_mb": 0,
                    "pids": [],
                },
            )
            item["count"] += 1
            item["pids"].append(pid)
            mem = proc.info.get("memory_info")
            if mem:
                item["memory_mb"] += int(mem.rss / (1024 * 1024))
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            continue

    items = list(entries.values())
    items.sort(key=lambda item: (-item["memory_mb"], item["name"].lower()))
    return items[:limit]


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

if _AUDIO_API_AVAILABLE:
    CLSID_MMDeviceEnumerator = GUID("{BCDE0395-E52F-467C-8E3D-C4579291692E}")
    IID_IAudioEndpointVolume = GUID("{5CDF2C82-841E-4546-9722-0CF74078229A}")
    IID_IPropertyStore = GUID("{886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99}")
    E_RENDER = 0
    E_MULTIMEDIA = 1
    STGM_READ = 0
    VT_LPWSTR = 31

    class IMMDevice(IUnknown):
        _iid_ = GUID("{D666063F-1587-4E43-81F1-B948E807363F}")
        _methods_ = [
            COMMETHOD(
                [],
                HRESULT,
                "Activate",
                (["in"], POINTER(GUID), "iid"),
                (["in"], wintypes.DWORD, "dwClsCtx"),
                (["in"], ctypes.c_void_p, "pActivationParams"),
                (["out"], POINTER(ctypes.c_void_p), "ppInterface"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "OpenPropertyStore",
                (["in"], wintypes.DWORD, "stgmAccess"),
                (["out"], POINTER(ctypes.c_void_p), "ppProperties"),
            ),
        ]

    class IMMDeviceEnumerator(IUnknown):
        _iid_ = GUID("{A95664D2-9614-4F35-A746-DE8DB63617E6}")
        _methods_ = [
            COMMETHOD(
                [],
                HRESULT,
                "EnumAudioEndpoints",
                (["in"], wintypes.DWORD, "dataFlow"),
                (["in"], wintypes.DWORD, "dwStateMask"),
                (["out"], POINTER(ctypes.c_void_p), "ppDevices"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetDefaultAudioEndpoint",
                (["in"], wintypes.DWORD, "dataFlow"),
                (["in"], wintypes.DWORD, "role"),
                (["out"], POINTER(POINTER(IMMDevice)), "ppEndpoint"),
            ),
        ]

    class IAudioEndpointVolume(IUnknown):
        _iid_ = IID_IAudioEndpointVolume
        _methods_ = [
            COMMETHOD([], HRESULT, "RegisterControlChangeNotify", (["in"], ctypes.c_void_p, "pNotify")),
            COMMETHOD([], HRESULT, "UnregisterControlChangeNotify", (["in"], ctypes.c_void_p, "pNotify")),
            COMMETHOD([], HRESULT, "GetChannelCount", (["out"], POINTER(wintypes.UINT), "pnChannelCount")),
            COMMETHOD([], HRESULT, "SetMasterVolumeLevel", (["in"], c_float, "fLevelDB"), (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "SetMasterVolumeLevelScalar", (["in"], c_float, "fLevel"), (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "GetMasterVolumeLevel", (["out"], POINTER(c_float), "pfLevelDB")),
            COMMETHOD([], HRESULT, "GetMasterVolumeLevelScalar", (["out"], POINTER(c_float), "pfLevel")),
            COMMETHOD([], HRESULT, "SetChannelVolumeLevel", (["in"], wintypes.UINT, "nChannel"), (["in"], c_float, "fLevelDB"), (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "SetChannelVolumeLevelScalar", (["in"], wintypes.UINT, "nChannel"), (["in"], c_float, "fLevel"), (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "GetChannelVolumeLevel", (["in"], wintypes.UINT, "nChannel"), (["out"], POINTER(c_float), "pfLevelDB")),
            COMMETHOD([], HRESULT, "GetChannelVolumeLevelScalar", (["in"], wintypes.UINT, "nChannel"), (["out"], POINTER(c_float), "pfLevel")),
            COMMETHOD([], HRESULT, "SetMute", (["in"], wintypes.BOOL, "bMute"), (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "GetMute", (["out"], POINTER(wintypes.BOOL), "pbMute")),
            COMMETHOD([], HRESULT, "GetVolumeStepInfo", (["out"], POINTER(wintypes.UINT), "pnStep"), (["out"], POINTER(wintypes.UINT), "pnStepCount")),
            COMMETHOD([], HRESULT, "VolumeStepUp", (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "VolumeStepDown", (["in"], ctypes.c_void_p, "pguidEventContext")),
            COMMETHOD([], HRESULT, "QueryHardwareSupport", (["out"], POINTER(wintypes.DWORD), "pdwHardwareSupportMask")),
            COMMETHOD([], HRESULT, "GetVolumeRange", (["out"], POINTER(c_float), "pflVolumeMindB"), (["out"], POINTER(c_float), "pflVolumeMaxdB"), (["out"], POINTER(c_float), "pflVolumeIncrementdB")),
        ]

    class PROPERTYKEY(ctypes.Structure):
        _fields_ = [("fmtid", GUID), ("pid", wintypes.DWORD)]

    class PROPVARIANT_UNION(ctypes.Union):
        _fields_ = [("pwszVal", wintypes.LPWSTR), ("ulVal", wintypes.ULONG)]

    class PROPVARIANT(ctypes.Structure):
        _anonymous_ = ("u",)
        _fields_ = [
            ("vt", wintypes.USHORT),
            ("wReserved1", wintypes.USHORT),
            ("wReserved2", wintypes.USHORT),
            ("wReserved3", wintypes.USHORT),
            ("u", PROPVARIANT_UNION),
        ]

    class IPropertyStore(IUnknown):
        _iid_ = IID_IPropertyStore
        _methods_ = [
            COMMETHOD([], HRESULT, "GetCount", (["out"], POINTER(wintypes.DWORD), "cProps")),
            COMMETHOD([], HRESULT, "GetAt", (["in"], wintypes.DWORD, "iProp"), (["out"], POINTER(PROPERTYKEY), "pkey")),
            COMMETHOD([], HRESULT, "GetValue", (["in"], POINTER(PROPERTYKEY), "key"), (["out"], POINTER(PROPVARIANT), "pv")),
            COMMETHOD([], HRESULT, "SetValue", (["in"], POINTER(PROPERTYKEY), "key"), (["in"], POINTER(PROPVARIANT), "propvar")),
            COMMETHOD([], HRESULT, "Commit"),
        ]

    PKEY_Device_FriendlyName = PROPERTYKEY(
        GUID("{A45C254E-DF1C-4EFD-8020-67D146A850E0}"), 14
    )
    _PropVariantClear = ctypes.windll.ole32.PropVariantClear
    _PropVariantClear.argtypes = [POINTER(PROPVARIANT)]
    _PropVariantClear.restype = HRESULT


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
        self._transparent_key = "#000001"
        self.configure(fg_color=self._transparent_key)

        self.pill_frame = ctk.CTkFrame(
            self,
            border_width=1,
            border_color="#3b404b",
            corner_radius=22,
            bg_color=self._transparent_key,
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
                "right": ["taskmanager", "sys_monitor", "clock", "background_apps", "tray", "volume_control"],
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

        self._popup_gap = 8
        self._popup_large_gap = 12
        self._tooltip_delay_ms = 240
        self._stack_menu_open_delay_ms = 90
        self._stack_menu_hide_delay_ms = 90
        self._fade_steps = 1
        self._fade_step_ms = 1
        self._popup_names = (
            "start_menu",
            "search_window",
            "tray_menu",
            "background_apps_menu",
            "background_app_context_menu",
            "volume_menu",
            "overflow_menu",
        )
        self.theme = {
            "bar_bg": self.config.get("bg_color", "#1e1e1e"),
            "bar_border": "#3b404b",
            "surface": "#16191f",
            "surface_alt": "#1c2027",
            "surface_muted": "#232833",
            "surface_hover": "#2b3340",
            "surface_active": "#173045",
            "surface_active_border": "#67b8ff",
            "surface_danger": "#3a2026",
            "surface_danger_hover": "#4a262d",
            "separator": "#2c313c",
            "text": "#f3f5f7",
            "text_dim": "#c7cdd6",
            "text_muted": "#8f98a4",
            "text_faint": "#68717d",
            "accent": "#3f9ef8",
            "accent_hover": "#2286e6",
            "danger": "#c64d56",
            "danger_hover": "#a93a43",
            "warning": "#f0b04f",
            "tooltip_bg": "#20242b",
            "tooltip_border": "#39404c",
        }

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
            self._transparent_key,
            "-topmost",
            True,
            "-alpha",
            float(self.config.get("opacity", 1.0)),
        )

        # ── Design tokens ──────────────────────────────────────────────────────
        _bg = self.theme["bar_bg"]
        _bh = max(26, h - 12)  # button height: 33px at default h=45
        _icr = 8  # corner radius, all buttons
        _icf = ("Segoe MDL2 Assets", 13)  # MDL2 icon font, all icon buttons
        _ich = self.theme["text_dim"]  # icon glyph colour
        _ihv = self.theme["surface_hover"]  # hover background
        self._btn_h = _bh  # stored so add_app_button can use it
        self._app_btn_w = _bh + 8  # app slots need a little extra width so icons don't clip
        self._button_corner_radius = _icr

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
            text_color=self.theme["text_muted"],
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
            fill=self.theme["text"],
            font=("Segoe UI Variable", 11, "bold"),
            anchor="center",
        )
        self._clock_date_id = self.clock_container.create_text(
            _cx,
            _cy + 8,
            text="",
            fill=self.theme["text_faint"],
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
            fill=self.theme["text_muted"],
            font=("Segoe UI Variable", 10),
            anchor="center",
        )
        self._sys_ram_id = self.sys_container.create_text(
            _sx,
            _sy + 8,
            text="RAM  0%",
            fill=self.theme["text_faint"],
            font=("Segoe UI Variable", 10),
            anchor="center",
        )

        # Separator between sys_monitor and clock — inserted by render_bar
        self._info_sep = ctk.CTkFrame(
            self.pill_frame, width=1, height=16, fg_color=self.theme["separator"]
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
        self.background_apps_btn = _ibtn(
            self.pill_frame, "\uec8f", self.toggle_background_apps
        )
        self.volume_control_btn = _ibtn(self.pill_frame, "\ue745", self.volume_control)
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
            fg_color=self.theme["accent"],
            hover_color=self.theme["accent_hover"],
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
            "background_apps": self.background_apps_btn,
            "tray": self.tray_btn,
            "volume_control": self.volume_control_btn,
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
        self.app_groups = {}
        self.hwnd_to_group = {}
        self._group_icon_cache = {}
        self._stack_menu_anchor = None
        self._stack_menu_group = None
        self._pinned_buttons = []
        self._pinned_separator = None
        self._volume_available = _AUDIO_API_AVAILABLE
        self._volume_slider_busy = False
        self._volume_poll_inflight = False
        self._volume_state_cache = (0.0, False)
        self._background_icon_cache = {}
        self._background_apps_request = 0

        psutil.cpu_percent(interval=None)
        # Pre-build the app search index off the main thread so first search is instant
        threading.Thread(target=_build_app_index, daemon=True).start()
        self.render_bar()
        self._schedule_volume_button_refresh()
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
            padx = 2 if key == "apps" else 4
            self.widget_map[key].pack(in_=zone, side="left", padx=padx)

    def _popup_target_alpha(self):
        return float(self.config.get("opacity", 1.0))

    def _popup_exists(self, attr):
        popup = getattr(self, attr, None)
        return popup is not None and popup.winfo_exists()

    def _destroy_popup(self, attr):
        popup = getattr(self, attr, None)
        if popup is not None:
            try:
                if popup.winfo_exists():
                    popup.destroy()
            except Exception:
                pass
        setattr(self, attr, None)

    def _close_all_popups(self, exclude=()):
        excluded = set(exclude)
        for attr in self._popup_names:
            if attr not in excluded:
                self._destroy_popup(attr)
        if "stack_menu" not in excluded:
            self._hide_stack_menu()
        if "tooltip" not in excluded:
            self._hide_tooltip()

    def _is_focus_in_widget(self, focused, widget):
        return (
            focused is not None
            and widget is not None
            and widget.winfo_exists()
            and str(focused).startswith(str(widget))
        )

    def _is_pointer_over_widget(self, widget):
        if widget is None or not widget.winfo_exists():
            return False
        mx, my = self.winfo_pointerxy()
        wx = widget.winfo_rootx()
        wy = widget.winfo_rooty()
        return wx <= mx <= wx + widget.winfo_width() and wy <= my <= wy + widget.winfo_height()

    def _popup_anchor_y(self, anchor_y, anchor_h, menu_h, gap=None):
        gap = self._popup_gap if gap is None else gap
        if self.config.get("position", "Top") == "Bottom":
            return anchor_y - menu_h - gap
        return anchor_y + anchor_h + gap

    def _create_popup_shell(
        self,
        attr,
        menu_w,
        menu_h,
        *,
        anchor_widget=None,
        align="left",
        gap=None,
        offset_x=0,
        x=None,
        y=None,
        corner_radius=20,
    ):
        self._destroy_popup(attr)
        popup = ctk.CTkToplevel(self)
        popup.overrideredirect(True)
        popup.attributes("-topmost", True)
        popup.configure(fg_color=self._transparent_key)
        popup.wm_attributes(
            "-transparentcolor",
            self._transparent_key,
            "-alpha",
            self._popup_target_alpha(),
        )

        if anchor_widget is None:
            anchor_x = self.winfo_x()
            anchor_w = self.winfo_width()
            anchor_y = self.winfo_y()
            anchor_h = self.winfo_height()
        else:
            anchor_x = anchor_widget.winfo_rootx()
            anchor_w = anchor_widget.winfo_width()
            anchor_y = anchor_widget.winfo_rooty()
            anchor_h = anchor_widget.winfo_height()

        if x is None:
            if align == "right":
                x = anchor_x + anchor_w - menu_w + offset_x
            elif align == "center":
                x = anchor_x + (anchor_w // 2) - (menu_w // 2) + offset_x
            else:
                x = anchor_x + offset_x
        if y is None:
            y = self._popup_anchor_y(anchor_y, anchor_h, menu_h, gap=gap)

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        x = max(10, min(int(x), screen_w - menu_w - 10))
        y = max(10, min(int(y), screen_h - menu_h - 10))
        popup.geometry(f"{menu_w}x{menu_h}+{x}+{y}")

        panel = ctk.CTkFrame(
            popup,
            corner_radius=corner_radius,
            fg_color=self.theme["surface"],
            border_width=1,
            border_color=self.theme["bar_border"],
        )
        panel.pack(fill="both", expand=True)
        setattr(self, attr, popup)
        return popup, panel

    def _bind_popup_focus_close(self, popup_attr, *, anchor_widget=None, allow=()):
        popup = getattr(self, popup_attr, None)
        if popup is None:
            return

        def _close_if_outside():
            if not self._popup_exists(popup_attr):
                return
            focused = self.focus_get()
            active_popup = getattr(self, popup_attr, None)
            if self._is_focus_in_widget(focused, active_popup):
                return
            for allowed_attr in allow:
                if self._is_focus_in_widget(focused, getattr(self, allowed_attr, None)):
                    return
            if anchor_widget is not None and self._is_pointer_over_widget(anchor_widget):
                return
            self._destroy_popup(popup_attr)

        popup.bind("<FocusOut>", lambda _event: popup.after(90, _close_if_outside))

    def _popup_header(self, parent, icon, title, subtitle=None, value=None):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(fill="x", padx=16, pady=(14, 6))

        icon_chip = ctk.CTkFrame(
            header,
            width=28,
            height=28,
            corner_radius=10,
            fg_color=self.theme["surface_alt"],
        )
        icon_chip.pack(side="left")
        icon_chip.pack_propagate(False)
        icon_label = ctk.CTkLabel(
            icon_chip,
            text=icon,
            font=("Segoe MDL2 Assets", 13),
            text_color=self.theme["text_dim"],
        )
        icon_label.pack(expand=True)

        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.pack(side="left", fill="x", expand=True, padx=(10, 0))
        ctk.CTkLabel(
            title_wrap,
            text=title,
            font=("Segoe UI Variable", 15, "bold"),
            text_color=self.theme["text"],
            anchor="w",
        ).pack(fill="x")
        if subtitle:
            ctk.CTkLabel(
                title_wrap,
                text=subtitle,
                font=("Segoe UI Variable", 11),
                text_color=self.theme["text_muted"],
                anchor="w",
            ).pack(fill="x")

        if value is not None:
            value_label = ctk.CTkLabel(
                header,
                text=value,
                font=("Segoe UI Variable", 12, "bold"),
                text_color=self.theme["text_muted"],
            )
            value_label.pack(side="right")
            return icon_label, value_label
        return icon_label, None

    def _popup_divider(self, parent, pady=(0, 10)):
        ctk.CTkFrame(parent, height=1, fg_color=self.theme["separator"]).pack(
            fill="x", padx=14, pady=pady
        )

    def _popup_button(self, parent, *, variant="ghost", **kwargs):
        base = {
            "ghost": {
                "fg_color": "transparent",
                "hover_color": self.theme["surface_hover"],
                "text_color": self.theme["text_dim"],
            },
            "soft": {
                "fg_color": self.theme["surface_alt"],
                "hover_color": self.theme["surface_hover"],
                "text_color": self.theme["text_dim"],
            },
            "primary": {
                "fg_color": self.theme["accent"],
                "hover_color": self.theme["accent_hover"],
                "text_color": "white",
            },
            "danger": {
                "fg_color": self.theme["danger"],
                "hover_color": self.theme["danger_hover"],
                "text_color": "white",
            },
            "danger_ghost": {
                "fg_color": "transparent",
                "hover_color": self.theme["surface_danger_hover"],
                "text_color": "#e3a4ab",
            },
        }[variant]
        return ctk.CTkButton(
            parent,
            corner_radius=10,
            border_width=0,
            font=("Segoe UI Variable", 11, "bold"),
            **base,
            **kwargs,
        )

    def _is_own_window(self, hwnd: int) -> bool:
        """Return True if hwnd belongs to WinBar itself or any of its popups."""
        get_ancestor = ctypes.windll.user32.GetAncestor
        if self._my_hwnd and hwnd == self._my_hwnd:
            return True
        for attr in self._popup_names:
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
        self._sync_volume_button_icon()
        self._schedule_volume_button_refresh()
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
            old_group = self.hwnd_to_group.get(self._active_hwnd)
            old_btn = self.active_app_buttons.get(old_group)
            if old_btn and old_btn.winfo_exists():
                self._set_group_button_state(old_btn, active=False)
        self._active_hwnd = hwnd
        # Highlight active app with accent border + tinted background
        new_group = self.hwnd_to_group.get(hwnd)
        new_btn = self.active_app_buttons.get(new_group)
        if new_btn and new_btn.winfo_exists():
            self._set_group_button_state(new_btn, active=True)

    def _group_label(self, group):
        count = len(group.get("windows", []))
        name = group.get("name", "App")
        return f"{name} ({count})" if count > 1 else name

    def _sync_group_icon(self, group_key, img):
        self._pending_icons.discard(group_key)
        if img:
            self._group_icon_cache[group_key] = img
        self._rebuild_group_buttons()

    def _handle_group_enter(self, widget, group_key):
        group = self.app_groups.get(group_key)
        if not group:
            return
        if len(group.get("windows", [])) > 1:
            self._hide_tooltip()
            self._schedule_stack_menu(widget, group_key)
        else:
            self._hide_stack_menu()
            self._schedule_tooltip(widget, self._group_label(group))

    def _handle_group_leave(self):
        self._hide_tooltip()
        self._schedule_stack_menu_hide()

    def _set_group_button_state(self, container, active=False):
        button = getattr(container, "_button", None)
        indicator = getattr(container, "_indicator", None)
        count = getattr(container, "_window_count", 1)
        if button and button.winfo_exists():
            if active:
                button.configure(
                    fg_color=self.theme["surface_active"],
                    border_width=1,
                    border_color=self.theme["surface_active_border"],
                )
            else:
                button.configure(fg_color=self.theme["surface_alt"], border_width=0)
        if indicator and indicator.winfo_exists():
            if active:
                indicator.configure(fg_color=self.theme["surface_active_border"])
            elif count > 1:
                indicator.configure(fg_color=self.theme["text_muted"])
            else:
                indicator.configure(fg_color="transparent")

    def _create_group_button(self, group_key, group):
        bh = self._btn_h
        bw = self._app_btn_w
        icon_size = int(bh * 0.55)
        img = self._group_icon_cache.get(group_key)
        icon = ctk.CTkImage(img, img, size=(icon_size, icon_size)) if img else None
        count = len(group.get("windows", []))
        container = ctk.CTkFrame(
            self.apps_container, fg_color="transparent", width=bw, height=bh
        )
        container.pack_propagate(False)
        btn = ctk.CTkButton(
            container,
            text="" if icon else group["name"][:8],
            image=icon,
            width=bw,
            height=bh,
            fg_color=self.theme["surface_alt"],
            hover_color=self.theme["surface_hover"],
            corner_radius=10,
            font=("Segoe UI Variable", 10),
            text_color=self.theme["text_dim"],
            anchor="center",
            border_spacing=0,
            command=lambda g=group_key: self.focus_group(g),
        )
        btn.pack(side="top")
        indicator = ctk.CTkFrame(
            container,
            width=12 if count > 1 else 8,
            height=2,
            corner_radius=2,
            fg_color=self.theme["text_muted"] if count > 1 else "transparent",
        )
        indicator.pack(side="top", pady=(2, 0))

        for widget in (container, btn, indicator):
            widget.bind(
                "<Button-2>", lambda e, g=group_key: self._close_group_window(g)
            )
            widget.bind(
                "<Enter>",
                lambda e, w=btn, g=group_key: self._handle_group_enter(w, g),
            )
            widget.bind("<Leave>", lambda e: self._handle_group_leave())

        container._full_name = self._group_label(group)
        container._is_pinned = False
        container._group_key = group_key
        container._button = btn
        container._indicator = indicator
        container._window_count = count
        self.active_app_buttons[group_key] = container
        self._set_group_button_state(
            container, active=self.hwnd_to_group.get(self._active_hwnd) == group_key
        )
        return container

    def _rebuild_group_buttons(self):
        if not hasattr(self, "active_app_buttons"):
            self.active_app_buttons = {}

        for btn in self.active_app_buttons.values():
            try:
                if btn.winfo_exists():
                    btn.destroy()
            except Exception:
                pass
        self.active_app_buttons = {}

        for group_key, group in self.app_groups.items():
            self._create_group_button(group_key, group)

        stale_icons = set(self._group_icon_cache) - set(self.app_groups)
        for group_key in stale_icons:
            self._group_icon_cache.pop(group_key, None)

        self._repack_app_buttons()
        self._update_active_indicator(self._active_hwnd)

    def _launch_path(self, path):
        try:
            os.startfile(path)
        except Exception:
            try:
                subprocess.Popen([path])
            except Exception:
                pass

    def _create_pinned_launcher(self, app):
        bh = self._btn_h
        bw = self._app_btn_w
        icon_size = int(bh * 0.55)
        path = app.get("path", "")
        exe = resolve_lnk(path) if path.lower().endswith(".lnk") else path
        img = get_icon_from_exe(exe) if exe and os.path.exists(exe) else None
        icon = ctk.CTkImage(img, img, size=(icon_size, icon_size)) if img else None
        btn = ctk.CTkButton(
            self.apps_container,
            text="" if icon else app.get("name", "App")[:3],
            image=icon,
            width=bw,
            height=bh,
            fg_color="transparent",
            hover_color=self.theme["surface_hover"],
            corner_radius=10,
            anchor="center",
            border_spacing=0,
            command=lambda p=path: self._launch_path(p),
        )
        btn._full_name = app.get("name", "App")
        btn._is_pinned = True
        btn.bind(
            "<Enter>",
            lambda e, w=btn, t=app.get("name", "App"): self._schedule_tooltip(w, t),
        )
        btn.bind("<Leave>", lambda e: self._hide_tooltip())
        return btn

    def _refresh_pinned_launchers(self):
        for btn in self._pinned_buttons:
            try:
                if btn.winfo_exists():
                    btn.destroy()
            except Exception:
                pass
        self._pinned_buttons = []

        if self._pinned_separator is not None:
            try:
                if self._pinned_separator.winfo_exists():
                    self._pinned_separator.destroy()
            except Exception:
                pass
            self._pinned_separator = None

        for app in self.config.get("pinned_apps", []):
            path = app.get("path")
            if path and os.path.exists(path):
                self._pinned_buttons.append(self._create_pinned_launcher(app))

        if self._pinned_buttons:
            self._pinned_separator = ctk.CTkFrame(
                self.apps_container,
                width=10,
                height=self._btn_h,
                fg_color="transparent",
            )
            ctk.CTkFrame(
                self._pinned_separator,
                width=1,
                height=max(16, self._btn_h - 10),
                fg_color=self.theme["text_faint"],
            ).pack(side="left", padx=(4, 0), pady=4)
            ctk.CTkFrame(
                self._pinned_separator,
                width=1,
                height=max(12, self._btn_h - 16),
                fg_color=self.theme["separator"],
            ).pack(side="left", padx=(2, 0), pady=6)

        self._repack_app_buttons()

    def _repack_app_buttons(self):
        if not hasattr(self, "active_app_buttons"):
            return

        item_padx = 1
        separator_padx = (4, 4)

        for child in self.apps_container.winfo_children():
            child.pack_forget()

        for btn in self._pinned_buttons:
            if btn.winfo_exists():
                btn.pack(side="left", padx=item_padx)

        running_buttons = [
            btn for btn in self.active_app_buttons.values() if btn.winfo_exists()
        ]

        if self._pinned_separator is not None and running_buttons:
            self._pinned_separator.pack(side="left", padx=separator_padx)

        MAX_VISIBLE = 14
        visible_count = 0
        self.overflow_apps.clear()

        for btn in running_buttons:
            if visible_count < MAX_VISIBLE:
                btn.pack(side="left", padx=item_padx)
                visible_count += 1
            else:
                self.overflow_apps.append(btn)

        if self.overflow_apps:
            self.overflow_btn.pack(side="left", padx=item_padx)
        elif self._popup_exists("overflow_menu"):
            self._destroy_popup("overflow_menu")

    def focus_group(self, group_key):
        if self.edit_mode:
            return
        group = self.app_groups.get(group_key)
        if not group:
            return

        active_in_group = (
            self._active_hwnd
            if self.hwnd_to_group.get(self._active_hwnd) == group_key
            else None
        )
        target = active_in_group or group["windows"][0]["hwnd"]
        self.focus_window(target)

    def _close_group_window(self, group_key):
        group = self.app_groups.get(group_key)
        if not group:
            return
        target = (
            self._active_hwnd
            if self.hwnd_to_group.get(self._active_hwnd) == group_key
            else group["windows"][0]["hwnd"]
        )
        self._close_window(target)

    def _show_stack_menu(self, widget, group_key):
        self._hide_stack_menu()
        group = self.app_groups.get(group_key)
        if not group or len(group.get("windows", [])) < 2:
            return

        self._stack_menu_anchor = widget
        self._stack_menu_group = group_key

        menu_h = min(len(group["windows"]) * 46 + 16, 320)
        menu_w = 300
        menu, frame = self._create_popup_shell(
            "_stack_menu",
            menu_w,
            menu_h,
            anchor_widget=widget,
            align="center",
            corner_radius=16,
        )
        scroll = ctk.CTkScrollableFrame(frame, fg_color="transparent", corner_radius=0)
        scroll.pack(fill="both", expand=True, padx=6, pady=6)

        windows = sorted(
            group["windows"],
            key=lambda item: item["hwnd"] != self._active_hwnd,
        )
        for item in windows:
            title = item["title"]
            if len(title) > 42:
                title = title[:41] + "\u2026"
            row = ctk.CTkButton(
                scroll,
                text=f"  {title}",
                image=self.active_app_buttons[group_key]._button.cget("image"),
                anchor="w",
                height=40,
                fg_color=(
                    self.theme["surface_alt"]
                    if item["hwnd"] == self._active_hwnd
                    else "transparent"
                ),
                hover_color=self.theme["surface_hover"],
                corner_radius=8,
                font=("Segoe UI Variable Semibold", 12),
                text_color=self.theme["text_dim"],
                command=lambda h=item["hwnd"]: [self.focus_window(h), self._hide_stack_menu()],
            )
            row.pack(fill="x", pady=2)

        menu.bind("<Leave>", lambda e: self._schedule_stack_menu_hide())
        menu.bind("<Enter>", lambda e: self._cancel_stack_menu_hide())
        self._fade_in(menu, self._popup_target_alpha())

    def _schedule_stack_menu(self, widget, group_key):
        self._cancel_stack_menu_hide()
        after_id = getattr(self, "_stack_menu_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
        self._stack_menu_after = self.after(
            self._stack_menu_open_delay_ms,
            lambda: self._show_stack_menu(widget, group_key),
        )

    def _cancel_stack_menu_hide(self):
        after_id = getattr(self, "_stack_menu_hide_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
            self._stack_menu_hide_after = None

    def _schedule_stack_menu_hide(self):
        self._cancel_stack_menu_hide()
        self._stack_menu_hide_after = self.after(
            self._stack_menu_hide_delay_ms, self._hide_stack_menu_if_outside
        )

    def _hide_stack_menu_if_outside(self):
        anchor = getattr(self, "_stack_menu_anchor", None)
        menu = getattr(self, "_stack_menu", None)
        try:
            px, py = self.winfo_pointerxy()
            inside_anchor = (
                anchor is not None
                and anchor.winfo_exists()
                and anchor.winfo_rootx() <= px <= anchor.winfo_rootx() + anchor.winfo_width()
                and anchor.winfo_rooty() <= py <= anchor.winfo_rooty() + anchor.winfo_height()
            )
            inside_menu = (
                menu is not None
                and menu.winfo_exists()
                and menu.winfo_rootx() <= px <= menu.winfo_rootx() + menu.winfo_width()
                and menu.winfo_rooty() <= py <= menu.winfo_rooty() + menu.winfo_height()
            )
            if not inside_anchor and not inside_menu:
                self._hide_stack_menu()
        except Exception:
            self._hide_stack_menu()

    def _hide_stack_menu(self):
        after_id = getattr(self, "_stack_menu_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
            self._stack_menu_after = None
        self._cancel_stack_menu_hide()
        menu = getattr(self, "_stack_menu", None)
        if menu is not None:
            try:
                menu.destroy()
            except Exception:
                pass
            self._stack_menu = None
        self._stack_menu_anchor = None
        self._stack_menu_group = None

    def _populate_start_pins(self):
        if not hasattr(self, "start_scroll") or not self.start_scroll.winfo_exists():
            return

        for child in self.start_scroll.winfo_children():
            child.destroy()

        pinned_apps = self.config.get("pinned_apps", [])
        if not pinned_apps:
            ctk.CTkLabel(
                self.start_scroll,
                text="\ue71d",
                font=("Segoe MDL2 Assets", 28),
                text_color=self.theme["text_faint"],
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.start_scroll,
                text="No pinned apps yet",
                font=("Segoe UI Variable", 13, "bold"),
                text_color=self.theme["text_dim"],
            ).pack()
            ctk.CTkLabel(
                self.start_scroll,
                text="Use Add App to pin an executable",
                font=("Segoe UI Variable", 11),
                text_color=self.theme["text_muted"],
            ).pack(pady=(2, 0))
            return

        for idx, app in enumerate(pinned_apps):
            if not os.path.exists(app.get("path", "")):
                continue

            row = ctk.CTkFrame(
                self.start_scroll,
                fg_color=self.theme["surface_alt"],
                corner_radius=12,
                border_width=1,
                border_color=self.theme["separator"],
            )
            row.pack(fill="x", pady=2, padx=5)

            exe = (
                resolve_lnk(app["path"])
                if app["path"].lower().endswith(".lnk")
                else app["path"]
            )
            img = get_icon_from_exe(exe) if exe else None
            icon = ctk.CTkImage(img, img, size=(32, 32)) if img else None

            self._popup_button(
                row,
                text=f"  {app['name']}",
                image=icon,
                anchor="w",
                height=48,
                font=("Segoe UI Variable Semibold", 14),
                text_color=self.theme["text"],
                command=lambda p=app["path"]: [
                    self._launch_path(p),
                    self._destroy_popup("start_menu"),
                ],
            ).pack(side="left", fill="x", expand=True, padx=(4, 0), pady=4)

            self._popup_button(
                row,
                text="\ue74d",
                width=36,
                height=36,
                font=("Segoe MDL2 Assets", 13),
                variant="danger_ghost",
                command=lambda i=idx: self.remove_pinned_app(i),
            ).pack(side="right", padx=(8, 8), pady=6)

    def add_pinned_app(self):
        path = filedialog.askopenfilename(
            title="Choose an app to pin",
            filetypes=[
                ("Applications", "*.exe *.lnk"),
                ("Executables", "*.exe"),
                ("Shortcuts", "*.lnk"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return

        normalized = os.path.normcase(os.path.normpath(path))
        existing = {
            os.path.normcase(os.path.normpath(app.get("path", "")))
            for app in self.config.get("pinned_apps", [])
        }
        if normalized in existing:
            if hasattr(self, "start_menu") and self.start_menu.winfo_exists():
                self.start_menu.focus_force()
            return

        name = os.path.splitext(os.path.basename(path))[0]
        self.config.setdefault("pinned_apps", []).append({"name": name, "path": path})
        self.save_layout()
        self._refresh_pinned_launchers()
        self._populate_start_pins()

        if hasattr(self, "start_menu") and self.start_menu.winfo_exists():
            self.start_menu.focus_force()

    def remove_pinned_app(self, index):
        pinned_apps = self.config.get("pinned_apps", [])
        if not (0 <= index < len(pinned_apps)):
            return

        pinned_apps.pop(index)
        self.save_layout()
        self._refresh_pinned_launchers()
        self._populate_start_pins()

    def update_open_apps(self):
        if not hasattr(self, "active_app_buttons"):
            self.active_app_buttons = {}
            self._refresh_pinned_launchers()

        current_windows = get_running_apps()
        cur_hwnds = {app["hwnd"] for app in current_windows}

        # Skip expensive repaint if the window set hasn't changed
        if cur_hwnds == self._prev_hwnds:
            self.after(500, self.update_open_apps)
            return
        self._prev_hwnds = cur_hwnds

        _skip = {
            "FloatingBar",
            "Limitens",
            "NVIDIA",
            "Overlay",
            "CTkToplevel",
            "Windows Shell Experience Host",
        }
        grouped = {}
        for app in current_windows:
            if any(x in app["title"] for x in _skip):
                continue
            hid = app["hwnd"]
            exe = get_exe_from_hwnd(hid)
            group_key = os.path.normcase(exe) if exe else f"title:{app['title'].lower()}"
            group_name = (
                os.path.splitext(os.path.basename(exe))[0] if exe else app["title"][:20]
            )
            group = grouped.setdefault(
                group_key,
                {"name": group_name, "exe": exe, "windows": []},
            )
            group["windows"].append({"hwnd": hid, "title": app["title"]})

            if (
                group_key not in self._group_icon_cache
                and group_key not in self._pending_icons
                and exe
            ):
                self._pending_icons.add(group_key)

                def _load(g=group_key, e=exe):
                    img = get_icon_from_exe(e)
                    self.after(0, lambda i=img, key=g: self._sync_group_icon(key, i))

                threading.Thread(target=_load, daemon=True).start()

        self.app_groups = grouped
        self.hwnd_to_group = {
            item["hwnd"]: group_key
            for group_key, group in grouped.items()
            for item in group["windows"]
        }
        self._rebuild_group_buttons()
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
                for attr in self._popup_names:
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
        """Open popups immediately so button feedback feels direct."""
        if target_alpha <= 0 or not window.winfo_exists():
            return
        window.wm_attributes("-alpha", target_alpha)

    def _show_tooltip(self, widget, text):
        self._hide_tooltip()
        tip = tk.Toplevel(self)
        tip.overrideredirect(True)
        tip.wm_attributes("-topmost", True)
        tip.configure(bg=self.theme["tooltip_bg"])

        # Plain tk widgets — synchronous creation, no CTkToplevel delayed init
        border = tk.Frame(tip, bg=self.theme["tooltip_border"], padx=1, pady=1)
        border.pack(fill="both", expand=True)
        inner = tk.Frame(border, bg=self.theme["tooltip_bg"], padx=10, pady=6)
        inner.pack(fill="both", expand=True)
        tk.Label(
            inner,
            text=text,
            font=("Segoe UI Variable", 11),
            fg=self.theme["text_dim"],
            bg=self.theme["tooltip_bg"],
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
            ty = widget.winfo_rooty() - th - self._popup_gap
        else:
            ty = widget.winfo_rooty() + widget.winfo_height() + self._popup_gap
        tip.geometry(f"+{tx}+{ty}")
        self._tooltip = tip

    def _schedule_tooltip(self, widget, text):
        self._cancel_tooltip()
        self._tooltip_after = self.after(
            self._tooltip_delay_ms, lambda: self._show_tooltip(widget, text)
        )

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
        if self._popup_exists("start_menu"):
            self._destroy_popup("start_menu")
            return
        self._close_all_popups(exclude=("start_menu",))

        menu_w = 300
        menu_h = 400
        self.start_menu, main_frame = self._create_popup_shell(
            "start_menu",
            menu_w,
            menu_h,
            anchor_widget=self.start_btn,
            align="left",
            corner_radius=18,
        )

        current_user = os.environ.get("USERNAME", "User")
        self._popup_header(
            main_frame,
            "\ue77b",
            current_user,
            subtitle="Pinned launchers and quick desktop controls",
        )
        self._popup_divider(main_frame, pady=(0, 8))

        pins_header = ctk.CTkFrame(main_frame, fg_color="transparent")
        pins_header.pack(fill="x", padx=15, pady=(0, 6))
        ctk.CTkLabel(
            pins_header,
            text="Pinned apps",
            font=("Segoe UI Variable", 12),
            text_color=self.theme["text_muted"],
        ).pack(side="left")
        self._popup_button(
            pins_header,
            text="+ Add App",
            height=30,
            width=88,
            variant="soft",
            command=self.add_pinned_app,
        ).pack(side="right")

        self.start_scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        self.start_scroll.pack(fill="both", expand=True, padx=10, pady=5)
        self._populate_start_pins()

        bottom_bar = ctk.CTkFrame(
            main_frame, fg_color=self.theme["surface_alt"], corner_radius=14
        )
        bottom_bar.pack(fill="x", side="bottom", padx=10, pady=10)

        settings_btn = self._popup_button(
            bottom_bar,
            text="\ue713  Settings",
            width=115,
            height=34,
            anchor="w",
            command=lambda: [
                os.system("start ms-settings:"),
                self._destroy_popup("start_menu"),
            ],
        )
        settings_btn.pack(side="left", padx=6, pady=6)
        self._popup_button(
            bottom_bar,
            text="\ue7e8",
            width=32,
            height=32,
            font=("Segoe MDL2 Assets", 14),
            variant="danger",
            command=lambda: os.system("shutdown /s /t 0"),
        ).pack(side="right", padx=6, pady=6)
        self._popup_button(
            bottom_bar,
            text="\uec46",
            width=32,
            height=32,
            font=("Segoe MDL2 Assets", 14),
            variant="soft",
            command=lambda: os.system(
                "rundll32.exe powrprof.dll,SetSuspendState 0,1,0"
            ),
        ).pack(side="right", padx=(0, 6), pady=6)

        self._bind_popup_focus_close("start_menu", anchor_widget=self.start_btn)
        self._fade_in(self.start_menu, self._popup_target_alpha())
        self.start_menu.focus_set()

    def toggle_search(self):
        if self.edit_mode:
            return
        if self._popup_exists("search_window"):
            self._destroy_popup("search_window")
            return
        self._close_all_popups(exclude=("search_window",))

        search_w, search_h = 500, 450
        self.search_window, main_frame = self._create_popup_shell(
            "search_window",
            search_w,
            search_h,
            align="center",
            gap=self._popup_large_gap,
            corner_radius=22,
        )
        self._popup_header(
            main_frame,
            "\ue721",
            "Search",
            subtitle="Launch apps and shortcuts from the Windows Start menu index",
        )
        search_area = ctk.CTkFrame(
            main_frame, fg_color=self.theme["surface_alt"], corner_radius=14
        )
        search_area.pack(fill="x", padx=15, pady=(15, 10))
        ctk.CTkLabel(
            search_area,
            text="\ue721",
            font=("Segoe MDL2 Assets", 14),
            text_color=self.theme["text_muted"],
        ).pack(side="left", padx=(12, 0))
        self.search_entry = ctk.CTkEntry(
            search_area,
            placeholder_text="Search apps...",
            width=400,
            height=45,
            font=("Segoe UI Variable", 16),
            fg_color="transparent",
            border_width=0,
            text_color=self.theme["text"],
            placeholder_text_color=self.theme["text_faint"],
        )
        self.search_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.search_entry.focus_force()
        self.results_scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        self.results_scroll.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.search_entry.bind("<KeyRelease>", self.update_search_results)
        self.search_window.bind("<Escape>", lambda e: self._destroy_popup("search_window"))
        self.search_entry.bind("<Return>", lambda e: self.launch_top_result())
        self._bind_popup_focus_close("search_window", anchor_widget=self.search_btn)
        self.update_search_results()
        self._fade_in(self.search_window, self._popup_target_alpha())
        self.search_window.focus_set()

    def update_search_results(self, event=None):
        query = self.search_entry.get()
        for child in self.results_scroll.winfo_children():
            child.destroy()
        if len(query) < 1:
            ctk.CTkLabel(
                self.results_scroll,
                text="\ue721",
                font=("Segoe MDL2 Assets", 30),
                text_color=self.theme["text_faint"],
            ).pack(pady=(54, 8))
            ctk.CTkLabel(
                self.results_scroll,
                text="Start typing to search your installed apps",
                font=("Segoe UI Variable", 13, "bold"),
                text_color=self.theme["text_dim"],
            ).pack()
            ctk.CTkLabel(
                self.results_scroll,
                text="Press Enter to launch the best match",
                font=("Segoe UI Variable", 11),
                text_color=self.theme["text_muted"],
            ).pack(pady=(4, 0))
            return
        apps = search_windows_apps(query)
        if not apps:
            ctk.CTkLabel(
                self.results_scroll,
                text="\ue721",
                font=("Segoe MDL2 Assets", 28),
                text_color=self.theme["text_faint"],
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.results_scroll,
                text="No results found",
                font=("Segoe UI Variable", 13),
                text_color=self.theme["text_muted"],
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
                corner_radius=10,
                fg_color="transparent",
                hover_color=self.theme["surface_hover"],
                font=("Segoe UI Variable Semibold", 13),
                text_color=self.theme["text_dim"],
                command=lambda p=app["path"]: [
                    os.startfile(p),
                    self._destroy_popup("search_window"),
                ],
            ).pack(fill="x", pady=2, padx=5)

    def launch_top_result(self):
        query = self.search_entry.get()
        if len(query) < 1:
            return
        apps = search_windows_apps(query)
        if apps:
            os.startfile(apps[0]["path"])
            self._destroy_popup("search_window")

    def _with_endpoint_volume(self, fn):
        if not _AUDIO_API_AVAILABLE:
            self._volume_available = False
            return None
        try:
            CoInitialize()
            enumerator = CoCreateInstance(
                CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL
            )
            device = enumerator.GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
            interface = device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, None)
            endpoint = cast(interface, POINTER(IAudioEndpointVolume))
            self._volume_available = True
            return fn(endpoint)
        except Exception:
            self._volume_available = False
            return None
        finally:
            if _AUDIO_API_AVAILABLE:
                try:
                    CoUninitialize()
                except Exception:
                    pass

    def _get_system_volume_state(self):
        def _read(endpoint):
            level = float(endpoint.GetMasterVolumeLevelScalar())
            mute = bool(endpoint.GetMute())
            return level, mute

        state = self._with_endpoint_volume(_read)
        return state if state is not None else (0.0, False)

    def _get_default_output_name(self):
        def _read_name(device):
            store_obj = device.OpenPropertyStore(STGM_READ)
            store = cast(store_obj, POINTER(IPropertyStore))
            value = store.GetValue(byref(PKEY_Device_FriendlyName))
            try:
                if value.vt == VT_LPWSTR and value.pwszVal:
                    return str(value.pwszVal)
                return "Default output"
            finally:
                _PropVariantClear(byref(value))

        if not _AUDIO_API_AVAILABLE:
            return "Audio unavailable"
        try:
            CoInitialize()
            enumerator = CoCreateInstance(
                CLSID_MMDeviceEnumerator, interface=IMMDeviceEnumerator, clsctx=CLSCTX_ALL
            )
            device = enumerator.GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
            name = _read_name(device)
            self._volume_available = True
            return name
        except Exception:
            self._volume_available = False
            return "Audio unavailable"
        finally:
            try:
                CoUninitialize()
            except Exception:
                pass

    def _set_system_volume(self, value):
        value = max(0.0, min(1.0, float(value)))
        return self._with_endpoint_volume(
            lambda endpoint: endpoint.SetMasterVolumeLevelScalar(value, None)
        )

    def _set_system_mute(self, muted):
        return self._with_endpoint_volume(
            lambda endpoint: endpoint.SetMute(bool(muted), None)
        )

    def _volume_glyph(self, level, muted):
        if muted or level <= 0.01:
            return "\ue74f"
        if level < 0.5:
            return "\ue993"
        return "\ue995"

    def _sync_volume_button_icon(self, state=None):
        if state is None:
            level, muted = self._volume_state_cache
        else:
            level, muted = state
            self._volume_state_cache = (level, muted)
        try:
            self.volume_control_btn.configure(
                text=self._volume_glyph(level, muted),
                text_color=self.theme["text_faint"] if muted else self.theme["text_dim"],
            )
        except Exception:
            pass

    def _schedule_volume_button_refresh(self):
        if self._volume_poll_inflight:
            return
        self._volume_poll_inflight = True

        def _worker():
            state = self._get_system_volume_state()

            def _apply():
                self._volume_poll_inflight = False
                self._sync_volume_button_icon(state)

            self.after(0, _apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_volume_menu_state(self, level, muted, device_name):
        if not self._popup_exists("volume_menu"):
            return

        self._sync_volume_button_icon((level, muted))
        percent = int(round(level * 100))
        glyph = self._volume_glyph(level, muted)

        self._volume_slider_busy = True
        try:
            if hasattr(self, "volume_slider"):
                self.volume_slider.set(percent)
        finally:
            self._volume_slider_busy = False

        if hasattr(self, "volume_status_label"):
            self.volume_status_label.configure(text="Muted" if muted else f"{percent}%")
        if hasattr(self, "volume_icon_label"):
            self.volume_icon_label.configure(
                text=glyph,
                text_color=self.theme["text_muted"] if muted else self.theme["text_dim"],
            )
        if hasattr(self, "volume_mute_btn"):
            self.volume_mute_btn.configure(
                text="Unmute" if muted else "Mute",
                fg_color=self.theme["surface_hover"] if muted else self.theme["surface_alt"],
            )
        if hasattr(self, "volume_hint_label"):
            self.volume_hint_label.configure(
                text="Audio unavailable"
                if not self._volume_available
                else ("Muted" if muted else "Master output")
            )
        if hasattr(self, "volume_device_label"):
            self.volume_device_label.configure(text=device_name)

    def _refresh_volume_menu_async(self):
        if not self._popup_exists("volume_menu"):
            return

        if hasattr(self, "volume_status_label"):
            self.volume_status_label.configure(text="...")
        if hasattr(self, "volume_device_label"):
            self.volume_device_label.configure(text="Loading audio device...")

        def _worker():
            level, muted = self._get_system_volume_state()
            device_name = self._get_default_output_name()
            self.after(
                0, lambda: self._apply_volume_menu_state(level, muted, device_name)
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _refresh_volume_menu(self):
        if not hasattr(self, "volume_menu") or not self.volume_menu.winfo_exists():
            return

        level, muted = self._get_system_volume_state()
        self._apply_volume_menu_state(level, muted, self._get_default_output_name())

    def _on_volume_slider(self, value):
        if self._volume_slider_busy:
            return
        scalar = max(0.0, min(1.0, float(value) / 100.0))
        self._set_system_volume(scalar)
        if scalar > 0:
            self._set_system_mute(False)
        self._refresh_volume_menu_async()

    def _toggle_system_mute(self):
        level, muted = self._volume_state_cache
        if level == 0.0 and not muted:
            level, muted = self._get_system_volume_state()
        self._set_system_mute(not muted)
        self._refresh_volume_menu_async()

    def _set_quick_volume(self, percent):
        self._on_volume_slider(percent)

    def _open_file_location(self, path):
        try:
            subprocess.Popen(["explorer", f"/select,{path}"])
        except Exception:
            pass

    def _get_cached_background_icon(self, path):
        norm = os.path.normcase(path)
        if norm not in self._background_icon_cache:
            self._background_icon_cache[norm] = get_icon_from_exe(path)
        img = self._background_icon_cache.get(norm)
        return ctk.CTkImage(img, img, size=(20, 20)) if img else None

    def _render_background_apps_loading(self, message="Loading background apps..."):
        grid = getattr(self, "background_apps_grid", None)
        if grid is None or not grid.winfo_exists():
            return
        for child in grid.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            grid,
            text="\ue895",
            font=("Segoe MDL2 Assets", 26),
            text_color=self.theme["text_faint"],
        ).pack(pady=(34, 6))
        ctk.CTkLabel(
            grid,
            text=message,
            font=("Segoe UI Variable", 12),
            text_color=self.theme["text_muted"],
        ).pack()

    def _apply_background_tile_icon(self, tile, path, request_id, img):
        if (
            request_id != self._background_apps_request
            or not self._popup_exists("background_apps_menu")
            or not tile.winfo_exists()
        ):
            return
        norm = os.path.normcase(path)
        self._background_icon_cache[norm] = img
        if img:
            icon = ctk.CTkImage(img, img, size=(20, 20))
            tile.configure(image=icon, text="")
            tile._icon_ref = icon

    def _load_background_tile_icon_async(self, tile, path, request_id):
        norm = os.path.normcase(path)
        cached = self._background_icon_cache.get(norm)
        if cached is not None:
            self._apply_background_tile_icon(tile, path, request_id, cached)
            return

        def _worker():
            img = get_icon_from_exe(path)
            self.after(
                0, lambda: self._apply_background_tile_icon(tile, path, request_id, img)
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _render_background_apps(self, request_id, bg_apps):
        if request_id != self._background_apps_request or not self._popup_exists(
            "background_apps_menu"
        ):
            return

        grid = getattr(self, "background_apps_grid", None)
        if grid is None or not grid.winfo_exists():
            return

        for child in grid.winfo_children():
            child.destroy()

        if not bg_apps:
            ctk.CTkLabel(
                grid,
                text="No background apps found",
                font=("Segoe UI Variable", 12),
                text_color=self.theme["text_muted"],
            ).pack(anchor="w", padx=6, pady=(10, 2))
            return

        for col in range(5):
            grid.grid_columnconfigure(col, weight=1)
        for idx, app in enumerate(bg_apps):
            tile = ctk.CTkButton(
                grid,
                text=(app["name"][:1] or "?").upper(),
                width=42,
                height=42,
                fg_color=self.theme["surface_alt"],
                hover_color=self.theme["surface_hover"],
                corner_radius=10,
                font=("Segoe UI Variable", 12, "bold"),
                text_color=self.theme["text_muted"],
                anchor="center",
                border_spacing=0,
                command=lambda p=app["path"]: self._launch_background_app(p),
            )
            tile.grid(
                row=idx // 5,
                column=idx % 5,
                padx=4,
                pady=4,
                sticky="nsew",
            )
            tooltip = app["name"]
            if app["count"] > 1:
                tooltip += f" ({app['count']})"
            tile.bind("<Enter>", lambda e, w=tile, t=tooltip: self._schedule_tooltip(w, t))
            tile.bind("<Leave>", lambda e: self._hide_tooltip())
            tile.bind("<ButtonPress-1>", lambda e: self._hide_tooltip(), add="+")
            tile.bind(
                "<Button-3>",
                lambda e, w=tile, a=app: self._show_background_app_menu(w, a),
                add="+",
            )
            self._load_background_tile_icon_async(tile, app["path"], request_id)

    def _load_background_apps_async(self):
        self._background_apps_request += 1
        request_id = self._background_apps_request

        def _worker():
            apps = get_background_apps(limit=20)
            self.after(0, lambda: self._render_background_apps(request_id, apps))

        threading.Thread(target=_worker, daemon=True).start()

    def _launch_background_app(self, path):
        self._hide_tooltip()
        self._launch_path(path)
        self._destroy_popup("background_apps_menu")

    def _refresh_background_apps_menu(self):
        if self._popup_exists("background_apps_menu"):
            self._destroy_popup("background_apps_menu")
            self.toggle_background_apps()

    def _end_background_app(self, app):
        self._hide_tooltip()
        for pid in app.get("pids", []):
            try:
                psutil.Process(pid).terminate()
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue
        self.after(180, self._refresh_background_apps_menu)

    def _show_background_app_menu(self, widget, app):
        self._hide_tooltip()
        menu, frame = self._create_popup_shell(
            "background_app_context_menu",
            176,
            126,
            anchor_widget=widget,
            align="center",
            corner_radius=14,
        )

        self._popup_button(
            frame,
            text="Open app",
            height=34,
            anchor="w",
            command=lambda p=app["path"]: [self._launch_background_app(p), menu.destroy()],
        ).pack(fill="x", padx=6, pady=(6, 2))
        self._popup_button(
            frame,
            text="Open file location",
            height=34,
            anchor="w",
            command=lambda p=app["path"]: [self._open_file_location(p), menu.destroy()],
        ).pack(fill="x", padx=6, pady=2)
        self._popup_button(
            frame,
            text="End task",
            height=34,
            anchor="w",
            variant="danger_ghost",
            command=lambda a=app: [self._end_background_app(a), menu.destroy()],
        ).pack(fill="x", padx=6, pady=(2, 6))

        self._bind_popup_focus_close(
            "background_app_context_menu",
            anchor_widget=widget,
            allow=("background_apps_menu",),
        )
        self._fade_in(menu, self._popup_target_alpha())
        menu.focus_set()

    def volume_control(self):
        if self.edit_mode:
            return
        if self._popup_exists("volume_menu"):
            self._destroy_popup("volume_menu")
            return
        self._close_all_popups(exclude=("volume_menu",))

        menu_w, menu_h = 280, 250
        self.volume_menu, f = self._create_popup_shell(
            "volume_menu",
            menu_w,
            menu_h,
            anchor_widget=self.volume_control_btn,
            align="right",
            corner_radius=20,
        )

        self.volume_icon_label, self.volume_status_label = self._popup_header(
            f,
            "\ue995",
            "Volume",
            subtitle="Master output level",
            value="0%",
        )
        self.volume_device_label = ctk.CTkLabel(
            f,
            text="Default output",
            font=("Segoe UI Variable", 11),
            text_color=self.theme["text_muted"],
            anchor="w",
        )
        self.volume_device_label.pack(fill="x", padx=18, pady=(0, 8))
        self._popup_divider(f, pady=(0, 10))

        slider_row = ctk.CTkFrame(f, fg_color="transparent")
        slider_row.pack(fill="x", padx=18, pady=(4, 12))
        ctk.CTkLabel(
            slider_row,
            text="\ue995",
            font=("Segoe MDL2 Assets", 15),
            text_color=self.theme["text_muted"],
        ).pack(side="left", padx=(0, 10))
        self.volume_slider = ctk.CTkSlider(
            slider_row,
            from_=0,
            to=100,
            number_of_steps=100,
            command=self._on_volume_slider,
            button_color=self.theme["accent"],
            button_hover_color=self.theme["accent_hover"],
            progress_color=self.theme["accent"],
        )
        self.volume_slider.pack(side="left", fill="x", expand=True)

        action_row = ctk.CTkFrame(f, fg_color="transparent")
        action_row.pack(fill="x", padx=18, pady=(0, 8))
        self.volume_mute_btn = self._popup_button(
            action_row,
            text="Mute",
            width=88,
            height=32,
            variant="soft",
            command=self._toggle_system_mute,
        )
        self.volume_mute_btn.pack(side="left")

        quick_row = ctk.CTkFrame(f, fg_color="transparent")
        quick_row.pack(fill="x", padx=18, pady=(0, 10))
        for idx in range(4):
            quick_row.grid_columnconfigure(idx, weight=1, uniform="volume_quick")
        for idx, pct in enumerate((25, 50, 75, 100)):
            btn = self._popup_button(
                quick_row,
                text=f"{pct}%",
                width=1,
                height=30,
                variant="ghost",
                command=lambda p=pct: self._set_quick_volume(p),
            )
            btn.grid(
                row=0,
                column=idx,
                sticky="ew",
                padx=(0, 4) if idx < 3 else (0, 0),
            )

        self.volume_hint_label = ctk.CTkLabel(
            f,
            text="Master output",
            font=("Segoe UI Variable", 11),
            text_color=self.theme["text_faint"],
        )
        self.volume_hint_label.pack(anchor="w", padx=18, pady=(2, 0))

        self._bind_popup_focus_close("volume_menu", anchor_widget=self.volume_control_btn)
        self._fade_in(self.volume_menu, self._popup_target_alpha())
        self.volume_menu.focus_set()
        self.after(1, self._refresh_volume_menu_async)

    def toggle_background_apps(self):
        if self.edit_mode:
            return
        if self._popup_exists("background_apps_menu"):
            self._destroy_popup("background_apps_menu")
            return
        self._close_all_popups(exclude=("background_apps_menu",))

        menu_w, menu_h = 268, 230
        self.background_apps_menu, panel = self._create_popup_shell(
            "background_apps_menu",
            menu_w,
            menu_h,
            anchor_widget=self.background_apps_btn,
            align="right",
            corner_radius=20,
        )
        self._popup_header(
            panel,
            "\uec8f",
            "Background apps",
            subtitle="Recently active processes without visible windows",
        )
        self._popup_divider(panel, pady=(0, 8))

        self.background_apps_grid = ctk.CTkScrollableFrame(
            panel, fg_color="transparent", corner_radius=0
        )
        self.background_apps_grid.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        self._render_background_apps_loading()
        self._bind_popup_focus_close(
            "background_apps_menu",
            anchor_widget=self.background_apps_btn,
            allow=("background_app_context_menu",),
        )
        self._fade_in(self.background_apps_menu, self._popup_target_alpha())
        self.background_apps_menu.focus_set()
        self.after(1, self._load_background_apps_async)

    def toggle_control_center(self):
        if self.edit_mode:
            return
        if self._popup_exists("tray_menu"):
            self._destroy_popup("tray_menu")
            return
        self._close_all_popups(exclude=("tray_menu",))

        menu_w, menu_h = 300, 420
        self.tray_menu, f = self._create_popup_shell(
            "tray_menu",
            menu_w,
            menu_h,
            anchor_widget=self.tray_btn,
            align="right",
            corner_radius=20,
        )
        self._popup_header(
            f,
            "\ue782",
            "WinBar",
            subtitle="Bar behavior, layout mode, and session controls",
        )
        self._popup_divider(f, pady=(0, 8))

        # Opacity
        ctk.CTkLabel(
            f,
            text="Opacity",
            font=("Segoe UI Variable", 11),
            text_color=self.theme["text_muted"],
        ).pack(anchor="w", padx=18)
        opacity_slider = ctk.CTkSlider(
            f,
            from_=0.1,
            to=1.0,
            number_of_steps=18,
            command=self.change_opacity,
            button_color=self.theme["accent"],
            button_hover_color=self.theme["accent_hover"],
            progress_color=self.theme["accent"],
        )
        opacity_slider.set(self._popup_target_alpha())
        opacity_slider.pack(pady=(2, 10), padx=18, fill="x")

        # Position
        ctk.CTkLabel(
            f,
            text="Position",
            font=("Segoe UI Variable", 11),
            text_color=self.theme["text_muted"],
        ).pack(anchor="w", padx=18)
        pos_menu = ctk.CTkOptionMenu(
            f,
            values=["Top", "Bottom"],
            command=self.change_position,
            width=214,
            fg_color=self.theme["surface_alt"],
            button_color=self.theme["surface_hover"],
            button_hover_color=self.theme["accent_hover"],
            dropdown_fg_color=self.theme["surface_alt"],
            dropdown_hover_color=self.theme["surface_hover"],
            text_color=self.theme["text_dim"],
        )
        pos_menu.set(self.config.get("position", "Top"))
        pos_menu.pack(pady=(2, 10), padx=18)

        # Actions
        self._popup_button(
            f,
            text="Edit Layout",
            height=32,
            variant="primary",
            command=self.toggle_edit_mode,
        ).pack(padx=18, pady=(0, 6), fill="x")
        self._popup_button(
            f,
            text="Quit WinBar",
            height=32,
            variant="danger",
            command=self.safe_exit,
        ).pack(padx=18, pady=(0, 14), fill="x")
        self._bind_popup_focus_close("tray_menu", anchor_widget=self.tray_btn)
        self._fade_in(self.tray_menu, self._popup_target_alpha())
        self.tray_menu.focus_set()

    def change_opacity(self, v):
        alpha = float(v)
        self.wm_attributes("-alpha", alpha)
        self.config["opacity"] = alpha
        for attr in self._popup_names:
            popup = getattr(self, attr, None)
            if popup is not None and popup.winfo_exists():
                try:
                    popup.wm_attributes("-alpha", alpha)
                except Exception:
                    pass
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
        self._close_all_popups()

    def toggle_overflow_menu(self):
        if self.edit_mode:
            return
        if self._popup_exists("overflow_menu"):
            self._destroy_popup("overflow_menu")
            return
        if not self.overflow_apps:
            return
        self._close_all_popups(exclude=("overflow_menu",))

        menu_h = min(len(self.overflow_apps) * 50 + 20, 500)
        menu_w = 220
        self.overflow_menu, main_frame = self._create_popup_shell(
            "overflow_menu",
            menu_w,
            menu_h,
            anchor_widget=self.overflow_btn,
            align="left",
            corner_radius=15,
        )
        self._popup_header(
            main_frame,
            "\ue712",
            "More apps",
            subtitle="Running apps that don't fit on the bar",
        )

        scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        for original_btn in self.overflow_apps:
            inner_btn = getattr(original_btn, "_button", original_btn)
            icon = inner_btn.cget("image")
            full_name = getattr(original_btn, "_full_name", "App")
            group_key = getattr(original_btn, "_group_key", None)
            btn = ctk.CTkButton(
                scroll,
                text=f"  {full_name[:25]}",
                image=icon,
                anchor="w",
                height=40,
                fg_color="transparent",
                hover_color=self.theme["surface_hover"],
                font=("Segoe UI Variable Semibold", 13),
                text_color=self.theme["text_dim"],
                command=lambda g=group_key: (
                    [self.focus_group(g), self._destroy_popup("overflow_menu")] if g else None
                ),
            )
            btn.pack(fill="x", pady=2)

        self._bind_popup_focus_close("overflow_menu", anchor_widget=self.overflow_btn)
        self._fade_in(self.overflow_menu, self._popup_target_alpha())
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
        self._close_all_popups()

        if self.edit_mode:
            self.pill_frame.configure(
                fg_color=self.config.get("bg_color"),
                border_color=self.theme["warning"],
                border_width=2,
            )
            self.active_window_label.configure(
                text="\ue70f  Edit Mode  —  drag to rearrange",
                text_color=self.theme["warning"],
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
                border_color=self.theme["bar_border"],
                border_width=1,
            )
            self.active_window_label.configure(text="", text_color=self.theme["text_muted"])
            self._set_cursor_recursive(self.pill_frame, "arrow")
            # Clear all widget borders
            for w in self.widget_map.values():
                try:
                    w.configure(border_width=0)
                except Exception:
                    pass
            self.edit_done_btn.pack_forget()
            # Restore active indicator if needed
            active_group = self.hwnd_to_group.get(self._active_hwnd)
            if active_group in self.active_app_buttons:
                btn = self.active_app_buttons[active_group]
                if btn.winfo_exists():
                    btn.configure(
                        fg_color=self.theme["surface_active"],
                        border_width=1,
                        border_color=self.theme["accent"],
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
