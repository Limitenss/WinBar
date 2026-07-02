import os, json, datetime, sys, subprocess, time, queue
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
    return f"{now.strftime('%#I:%M %p')}  |  {now.strftime('%a, %b %#d')}"


def trim_log_file(log_path, max_bytes=256_000):
    try:
        if not os.path.exists(log_path):
            return

        size = os.path.getsize(log_path)
        if size <= max_bytes:
            return

        with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
            data = f.read()

        trimmed = data[-max_bytes:]
        first_newline = trimmed.find("\n")
        if first_newline != -1:
            trimmed = trimmed[first_newline + 1 :]

        with open(log_path, "w", encoding="utf-8") as f:
            f.write(trimmed)
    except Exception:
        pass


def log_error(context, exc):
    try:
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        log_path = os.path.join(base_dir, "winbar.log")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {context}: {type(exc).__name__}: {exc}\n")

        trim_log_file(log_path)
    except Exception:
        pass


def log_event(context, message):
    try:
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        log_path = os.path.join(base_dir, "winbar.log")
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        with open(log_path, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {context}: {message}\n")

        trim_log_file(log_path)
    except Exception:
        pass


def validate_config(config, default_config):
    config["position"] = config.get("position", "Top")
    if config["position"] not in ("Top", "Bottom"):
        config["position"] = "Top"

    config["width_percent"] = max(40, min(100, int(config.get("width_percent", 95))))
    config["height"] = max(32, min(80, int(config.get("height", 45))))
    config["opacity"] = max(0.2, min(1.0, float(config.get("opacity", 0.631))))
    if config.get("bg_color") in (None, LEGACY_BAR_BG):
        config["bg_color"] = default_config.get("bg_color", DEFAULT_BAR_BG)

    if not isinstance(config.get("layout"), dict):
        config["layout"] = {
            "left": list(default_config["layout"].get("left", [])),
            "center": list(default_config["layout"].get("center", [])),
            "right": list(default_config["layout"].get("right", [])),
        }

    if not isinstance(config.get("pinned_apps"), list):
        config["pinned_apps"] = []

    hover_tooltips = config.get("hover_tooltips", True)
    if isinstance(hover_tooltips, str):
        config["hover_tooltips"] = hover_tooltips.strip().lower() not in (
            "0",
            "false",
            "no",
            "off",
        )
    else:
        config["hover_tooltips"] = bool(hover_tooltips)

    return config


def load_config(config_path, default_config):
    if not os.path.exists(config_path):
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(default_config, f, indent=4)

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)
    except Exception:
        loaded = {}

    config = default_config | loaded
    return validate_config(config, default_config)


_app_index: list = []
_app_index_time: float = 0.0
_app_index_ready = False
_app_index_building = False
_app_index_lock = threading.Lock()
_APP_INDEX_TTL = 60.0  # rebuild every 60 seconds
LEGACY_BAR_BG = "#1e1e1e"
DEFAULT_BAR_BG = "#101216"
DEFAULT_BAR_BORDER = "#343a46"
DEFAULT_BAR_RADIUS = 22
TRANSPARENT_KEY = "#000001"
WINDOW_TITLE_BLACKLIST = [
    "Program Manager",
    "Microsoft Text Input Application",
    "NVIDIA GeForce Overlay",
    "Discord Updater",
    "CTkToplevel",
]


def _build_app_index():
    global _app_index, _app_index_time, _app_index_ready, _app_index_building
    search_paths = [
        os.path.join(
            os.environ.get("ProgramData", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
        ),
        os.path.join(
            os.environ.get("AppData", ""),
            "Microsoft",
            "Windows",
            "Start Menu",
            "Programs",
        ),
    ]
    entries = []
    try:
        for base in search_paths:
            if os.path.exists(base):
                for root, _, files in os.walk(base):
                    for f in files:
                        if f.lower().endswith(".lnk"):
                            name = f[:-4]
                            lower_name = name.lower().strip()
                            entries.append(
                                {
                                    "name": name,
                                    "path": os.path.join(root, f),
                                    "_lower_name": lower_name,
                                    "_words": tuple(lower_name.split()),
                                }
                            )
        entries.sort(key=lambda app: app["_lower_name"])
        with _app_index_lock:
            _app_index = entries
            _app_index_time = time.monotonic()
            _app_index_ready = True
    finally:
        with _app_index_lock:
            _app_index_building = False


def _start_app_index_build(force=False):
    global _app_index_building
    with _app_index_lock:
        if _app_index_building:
            return False
        fresh = _app_index and time.monotonic() - _app_index_time <= _APP_INDEX_TTL
        if fresh and not force:
            return False
        _app_index_building = True
    threading.Thread(target=_build_app_index, daemon=True).start()
    return True


def score_app_match(app_name, query):
    name = app_name.lower().strip()
    q = query.lower().strip()

    if not q:
        return -1

    if name == q:
        return 300
    if name.startswith(q):
        return 200
    if q in name:
        return 100

    words = name.split()
    if any(word.startswith(q) for word in words):
        return 80

    return -1


def search_windows_apps(query):
    global _app_index, _app_index_time
    with _app_index_lock:
        index = list(_app_index)
        index_time = _app_index_time
    if not index:
        _start_app_index_build()
        return []
    if time.monotonic() - index_time > _APP_INDEX_TTL:
        _start_app_index_build(force=True)

    q = query.lower().strip()
    if not q:
        return []

    ranked = []
    for app in index:
        name = app.get("_lower_name") or app["name"].lower().strip()
        if name == q:
            score = 300
        elif name.startswith(q):
            score = 200
        elif q in name:
            score = 100
        elif any(word.startswith(q) for word in app.get("_words", name.split())):
            score = 80
        else:
            score = -1
        if score >= 0:
            ranked.append((score, name, app))

    ranked.sort(key=lambda item: (-item[0], item[1]))
    return [item[2] for item in ranked[:7]]


def normalize_app_path(path):
    return os.path.normcase(os.path.abspath(path)).strip()


def is_pinned_app_duplicate(pinned_apps, new_path):
    target = normalize_app_path(new_path)
    for app in pinned_apps:
        existing = app.get("path")
        if existing and normalize_app_path(existing) == target:
            return True
    return False


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


_SHGFI_ICON = 0x000000100
_SHGFI_LARGEICON = 0x000000000
_SHGFI_USEFILEATTRIBUTES = 0x000000010
_FILE_ATTRIBUTE_NORMAL = 0x80
_WM_GETICON = 0x007F
_ICON_SMALL = 0
_ICON_BIG = 1
_ICON_SMALL2 = 2
_SMTO_ABORTIFHUNG = 0x0002
_GCLP_HICON = -14
_GCLP_HICONSM = -34


class _SHFILEINFOW(ctypes.Structure):
    _fields_ = [
        ("hIcon", wintypes.HICON),
        ("iIcon", ctypes.c_int),
        ("dwAttributes", wintypes.DWORD),
        ("szDisplayName", ctypes.c_wchar * 260),
        ("szTypeName", ctypes.c_wchar * 80),
    ]


try:
    _SHGetFileInfoW = ctypes.windll.shell32.SHGetFileInfoW
    _SHGetFileInfoW.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        ctypes.POINTER(_SHFILEINFOW),
        wintypes.UINT,
        wintypes.UINT,
    ]
    _SHGetFileInfoW.restype = ctypes.c_size_t
except Exception:
    _SHGetFileInfoW = None

try:
    _SendMessageTimeoutW = ctypes.windll.user32.SendMessageTimeoutW
    _SendMessageTimeoutW.argtypes = [
        wintypes.HWND,
        wintypes.UINT,
        wintypes.WPARAM,
        wintypes.LPARAM,
        wintypes.UINT,
        wintypes.UINT,
        ctypes.POINTER(ctypes.c_size_t),
    ]
    _SendMessageTimeoutW.restype = wintypes.LPARAM
except Exception:
    _SendMessageTimeoutW = None

try:
    _GetClassLongPtrW = ctypes.windll.user32.GetClassLongPtrW
    _GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
    _GetClassLongPtrW.restype = ctypes.c_size_t
except Exception:
    try:
        _GetClassLongPtrW = ctypes.windll.user32.GetClassLongW
        _GetClassLongPtrW.argtypes = [wintypes.HWND, ctypes.c_int]
        _GetClassLongPtrW.restype = wintypes.DWORD
    except Exception:
        _GetClassLongPtrW = None


def _extract_hicon_shell(path):
    if _SHGetFileInfoW is None or not path:
        return None
    try:
        sfi = _SHFILEINFOW()
        flags = _SHGFI_ICON | _SHGFI_LARGEICON
        if not os.path.exists(path):
            flags |= _SHGFI_USEFILEATTRIBUTES
        res = _SHGetFileInfoW(
            path,
            _FILE_ATTRIBUTE_NORMAL,
            ctypes.byref(sfi),
            ctypes.sizeof(sfi),
            flags,
        )
        if res == 0 or not sfi.hIcon:
            return None
        return sfi.hIcon
    except Exception:
        return None


def _send_get_icon(hwnd, icon_size):
    if _SendMessageTimeoutW is None:
        return None
    try:
        result = ctypes.c_size_t()
        ok = _SendMessageTimeoutW(
            hwnd,
            _WM_GETICON,
            icon_size,
            0,
            _SMTO_ABORTIFHUNG,
            50,
            ctypes.byref(result),
        )
        if ok and result.value:
            return result.value
    except Exception:
        return None
    return None


def _get_class_icon(hwnd, index):
    if _GetClassLongPtrW is None:
        return None
    try:
        hicon = _GetClassLongPtrW(hwnd, index)
        return hicon or None
    except Exception:
        return None


def get_icon_from_hwnd(hwnd):
    if not hwnd:
        return None
    for hicon in (
        _send_get_icon(hwnd, _ICON_BIG),
        _send_get_icon(hwnd, _ICON_SMALL2),
        _send_get_icon(hwnd, _ICON_SMALL),
        _get_class_icon(hwnd, _GCLP_HICON),
        _get_class_icon(hwnd, _GCLP_HICONSM),
    ):
        if hicon:
            img = _hicon_to_pil(hicon)
            if img:
                return img
    return None


def _hicon_to_pil(hicon):
    hdc_screen = None
    hdc_mem = None
    hbmp = None
    try:
        hdc_screen = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
        hbmp = win32ui.CreateBitmap()
        hbmp.CreateCompatibleBitmap(hdc_screen, 32, 32)
        hdc_mem = hdc_screen.CreateCompatibleDC()
        hdc_mem.SelectObject(hbmp)
        hdc_mem.DrawIcon((0, 0), hicon)
        bmpstr = hbmp.GetBitmapBits(True)
        return Image.frombuffer("RGBA", (32, 32), bmpstr, "raw", "BGRA", 0, 1)
    except Exception:
        return None
    finally:
        try:
            if hdc_mem is not None:
                hdc_mem.DeleteDC()
        except Exception:
            pass
        try:
            if hdc_screen is not None:
                win32gui.ReleaseDC(0, hdc_screen.GetHandleOutput())
        except Exception:
            pass


def get_icon_from_exe(exe_path):
    if not exe_path:
        return None

    hicon = None
    extras = []
    try:
        large, small = win32gui.ExtractIconEx(exe_path, 0)
        if large:
            hicon = large[0]
            extras.extend(large[1:])
            extras.extend(small)
        elif small:
            hicon = small[0]
            extras.extend(small[1:])
    except Exception:
        hicon = None

    try:
        if hicon is None:
            hicon = _extract_hicon_shell(exe_path)

        if not hicon:
            return None

        return _hicon_to_pil(hicon)
    finally:
        for extra in extras:
            try:
                win32gui.DestroyIcon(extra)
            except Exception:
                pass
        if hicon:
            try:
                win32gui.DestroyIcon(hicon)
            except Exception:
                pass


ABM_NEW = 0x00000000
ABM_REMOVE = 0x00000001
ABM_SETPOS = 0x00000003
ABM_SETSTATE = 0x0000000A
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
            COMMETHOD(
                [],
                HRESULT,
                "RegisterControlChangeNotify",
                (["in"], ctypes.c_void_p, "pNotify"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "UnregisterControlChangeNotify",
                (["in"], ctypes.c_void_p, "pNotify"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetChannelCount",
                (["out"], POINTER(wintypes.UINT), "pnChannelCount"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetMasterVolumeLevel",
                (["in"], c_float, "fLevelDB"),
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetMasterVolumeLevelScalar",
                (["in"], c_float, "fLevel"),
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetMasterVolumeLevel",
                (["out"], POINTER(c_float), "pfLevelDB"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetMasterVolumeLevelScalar",
                (["out"], POINTER(c_float), "pfLevel"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetChannelVolumeLevel",
                (["in"], wintypes.UINT, "nChannel"),
                (["in"], c_float, "fLevelDB"),
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetChannelVolumeLevelScalar",
                (["in"], wintypes.UINT, "nChannel"),
                (["in"], c_float, "fLevel"),
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetChannelVolumeLevel",
                (["in"], wintypes.UINT, "nChannel"),
                (["out"], POINTER(c_float), "pfLevelDB"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetChannelVolumeLevelScalar",
                (["in"], wintypes.UINT, "nChannel"),
                (["out"], POINTER(c_float), "pfLevel"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetMute",
                (["in"], wintypes.BOOL, "bMute"),
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [], HRESULT, "GetMute", (["out"], POINTER(wintypes.BOOL), "pbMute")
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetVolumeStepInfo",
                (["out"], POINTER(wintypes.UINT), "pnStep"),
                (["out"], POINTER(wintypes.UINT), "pnStepCount"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "VolumeStepUp",
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "VolumeStepDown",
                (["in"], ctypes.c_void_p, "pguidEventContext"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "QueryHardwareSupport",
                (["out"], POINTER(wintypes.DWORD), "pdwHardwareSupportMask"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetVolumeRange",
                (["out"], POINTER(c_float), "pflVolumeMindB"),
                (["out"], POINTER(c_float), "pflVolumeMaxdB"),
                (["out"], POINTER(c_float), "pflVolumeIncrementdB"),
            ),
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
            COMMETHOD(
                [], HRESULT, "GetCount", (["out"], POINTER(wintypes.DWORD), "cProps")
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetAt",
                (["in"], wintypes.DWORD, "iProp"),
                (["out"], POINTER(PROPERTYKEY), "pkey"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "GetValue",
                (["in"], POINTER(PROPERTYKEY), "key"),
                (["out"], POINTER(PROPVARIANT), "pv"),
            ),
            COMMETHOD(
                [],
                HRESULT,
                "SetValue",
                (["in"], POINTER(PROPERTYKEY), "key"),
                (["in"], POINTER(PROPVARIANT), "propvar"),
            ),
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
tray_icon = None
tray_icon_lock = threading.Lock()


def register_appbar(window_id, bar_height, padding_y, edge=ABE_TOP):
    global global_abd
    global_abd = APPBARDATA()
    global_abd.cbSize = ctypes.sizeof(APPBARDATA)
    global_abd.hWnd = window_id
    global_abd.uEdge = edge
    ctypes.windll.shell32.SHAppBarMessage(ABM_NEW, ctypes.byref(global_abd))
    sw = ctypes.windll.user32.GetSystemMetrics(0)
    sh = ctypes.windll.user32.GetSystemMetrics(1)
    reserved_h = max(1, int(bar_height + padding_y))
    if edge == ABE_TOP:
        global_abd.rc.top, global_abd.rc.bottom = 0, reserved_h
    else:
        global_abd.rc.top, global_abd.rc.bottom = sh - reserved_h, sh
    global_abd.rc.left, global_abd.rc.right = 0, sw
    ctypes.windll.shell32.SHAppBarMessage(ABM_SETPOS, ctypes.byref(global_abd))


def unregister_appbar():
    global global_abd
    if global_abd:
        ctypes.windll.shell32.SHAppBarMessage(ABM_REMOVE, ctypes.byref(global_abd))
        global_abd = None


def get_taskbar_hwnds():
    hwnds = []
    primary = win32gui.FindWindow("Shell_TrayWnd", None)
    if primary:
        hwnds.append(primary)

    def _collect(hwnd, _param):
        try:
            if win32gui.GetClassName(hwnd) == "Shell_SecondaryTrayWnd":
                hwnds.append(hwnd)
        except Exception:
            pass

    try:
        win32gui.EnumWindows(_collect, None)
    except Exception:
        pass

    return list(dict.fromkeys(hwnds))


def is_taskbar_visible():
    try:
        return any(win32gui.IsWindowVisible(hwnd) for hwnd in get_taskbar_hwnds())
    except Exception:
        return False


def set_taskbar_visibility(visible=True):
    abd = APPBARDATA()
    abd.cbSize = ctypes.sizeof(abd)
    abd.hWnd = win32gui.FindWindow("Shell_TrayWnd", None)
    if visible:
        abd.lParam = 2
        ctypes.windll.shell32.SHAppBarMessage(ABM_SETSTATE, ctypes.byref(abd))
    cmd = win32con.SW_SHOW if visible else win32con.SW_HIDE
    for hwnd_tray in get_taskbar_hwnds():
        win32gui.ShowWindow(hwnd_tray, cmd)
    hwnd_start = win32gui.FindWindow("Button", None)
    if not hwnd_start:
        hwnd_start = win32gui.FindWindowEx(0, 0, "Button", None)
    if hwnd_start:
        win32gui.ShowWindow(hwnd_start, cmd)
    if visible:
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
    except Exception as exc:
        log_error("install_to_startup", exc)


def create_tray_image():
    image = Image.new("RGB", (64, 64), color=(16, 18, 22))
    draw = ImageDraw.Draw(image)
    draw.rounded_rectangle((8, 20, 56, 44), radius=12, fill=(35, 41, 52))
    draw.rounded_rectangle((12, 24, 52, 40), radius=8, fill=(77, 183, 255))
    draw.line((16, 25, 48, 25), fill=(153, 216, 255), width=1)
    return image


def stop_system_tray():
    global tray_icon
    with tray_icon_lock:
        icon = tray_icon
        tray_icon = None

    if icon is None:
        return

    try:
        icon.visible = False
    except Exception:
        pass

    try:
        icon.stop()
    except Exception as exc:
        log_error("stop_system_tray", exc)


def setup_system_tray(app_instance):
    global tray_icon

    def on_quit(icon, item):
        app_instance._post_ui(app_instance.safe_exit)

    icon = pystray.Icon(
        "FloatingBar",
        create_tray_image(),
        "Floating Bar",
        pystray.Menu(pystray.MenuItem("Quit", on_quit)),
    )
    with tray_icon_lock:
        tray_icon = icon

    try:
        icon.run()
    finally:
        with tray_icon_lock:
            if tray_icon is icon:
                tray_icon = None


class FloatingBar(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Limitens_FloatingBar")
        self.overrideredirect(True)
        self._transparent_key = TRANSPARENT_KEY
        self.configure(fg_color=self._transparent_key)

        self.pill_frame = ctk.CTkFrame(
            self,
            border_width=1,
            border_color=DEFAULT_BAR_BORDER,
            corner_radius=DEFAULT_BAR_RADIUS,
            bg_color=self._transparent_key,
        )
        self.pill_frame.pack(fill="both", expand=True)
        self.pill_frame.grid_rowconfigure(0, weight=1)

        DEFAULT_CONFIG = {
            "position": "Top",
            "width_percent": 95,
            "height": 45,
            "bg_color": DEFAULT_BAR_BG,
            "opacity": 0.631,
            "layout": {
                "left": ["start", "apps", "active_window"],
                "center": ["search"],
                "right": [
                    "taskmanager",
                    "sys_monitor",
                    "clock",
                    "background_apps",
                    "tray",
                    "volume_control",
                ],
            },
            "pinned_apps": [],
            "hover_tooltips": True,
        }

        if getattr(sys, "frozen", False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        self.config_path = os.path.join(application_path, "config.json")
        self.config = load_config(self.config_path, DEFAULT_CONFIG)

        self._tooltip_delay_ms = 220
        self._tooltip_hide_delay_ms = 90
        self._stack_menu_open_delay_ms = 90
        self._stack_menu_hide_delay_ms = 160
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
            "bar_bg": self.config.get("bg_color", DEFAULT_BAR_BG),
            "bar_border": DEFAULT_BAR_BORDER,
            "bar_highlight": "#49515f",
            "surface": "#12151b",
            "surface_alt": "#1a1f28",
            "surface_muted": "#232a36",
            "surface_hover": "#2d3544",
            "surface_active": "#172b3a",
            "surface_active_border": "#8fd0ff",
            "pinned_hover": "#203247",
            "pinned_press": "#18384f",
            "surface_danger": "#3b2026",
            "surface_danger_hover": "#502832",
            "separator": "#303642",
            "text": "#f5f7fa",
            "text_dim": "#d2d8e2",
            "text_muted": "#9da7b5",
            "text_faint": "#65707f",
            "accent": "#4db7ff",
            "accent_hover": "#329ee8",
            "danger": "#d75f6a",
            "danger_hover": "#b74752",
            "danger_text": "#f0a6ae",
            "warning": "#d7b56d",
            "warning_dark": "#665128",
            "warning_text": "#17120a",
            "drop_target": "#8fd0ff",
            "tooltip_bg": "#181c23",
            "tooltip_border": "#424a57",
        }
        self.fonts = {
            "bar_label": ("Segoe UI Variable", 11),
            "bar_label_bold": ("Segoe UI Variable", 11, "bold"),
            "bar_minor": ("Segoe UI Variable", 9),
            "bar_meter": ("Segoe UI Variable", 10),
            "caption": ("Segoe UI Variable", 11),
            "caption_bold": ("Segoe UI Variable", 11, "bold"),
            "label": ("Segoe UI Variable", 12),
            "label_bold": ("Segoe UI Variable", 12, "bold"),
            "body": ("Segoe UI Variable", 13),
            "body_bold": ("Segoe UI Variable", 13, "bold"),
            "title": ("Segoe UI Variable Display", 15, "bold"),
            "search": ("Segoe UI Variable Display", 16),
            "icon_sm": ("Segoe MDL2 Assets", 13),
            "icon_md": ("Segoe MDL2 Assets", 14),
            "icon_bar": ("Segoe MDL2 Assets", 15),
            "icon_lg": ("Segoe MDL2 Assets", 28),
            "icon_xl": ("Segoe MDL2 Assets", 30),
        }
        self._edge_gap = 5
        self.metrics = {
            "bar_radius": DEFAULT_BAR_RADIUS,
            "button_radius": 8,
            "indicator_radius": 1,
            "chip_radius": 10,
            "chip_size": 28,
            "tile_radius": 10,
            "row_radius": 10,
            "panel_radius": 14,
            "popup_radius": 20,
            "popup_radius_compact": 16,
            "tooltip_radius": 10,
            "popup_pad": 18,
            "popup_inner_pad": 10,
            "popup_gap": 8,
            "popup_large_gap": 12,
            "popup_header_pad_x": 16,
            "popup_header_pad_y": (14, 6),
            "divider_pad_x": 14,
            "section_pad_x": 15,
            "control_pad_x": 18,
            "toolbar_pad": 10,
            "widget_gap": 4,
            "tight_gap": 2,
            "control_gap": 6,
            "grid_gap": 4,
            "menu_button_h": 32,
            "menu_row_h": 40,
            "menu_large_row_h": 46,
            "tile_size": 42,
            "action_button_w": 88,
            "settings_button_w": 115,
            "option_menu_w": 214,
            "indicator_active_w": 14,
            "indicator_group_w": 10,
            "indicator_w": 12,
            "indicator_h": 2,
            "pinned_hover_w": 4,
            "pinned_anim_step_ms": 14,
            "pinned_anim_steps": 6,
            "pinned_separator_w": 12,
            "line_w": 1,
            "separator_h": 16,
            "bar_highlight_h": 1,
            "done_button_w": 56,
            "search_entry_h": 45,
            "search_entry_w": 400,
            "small_icon_button": 32,
            "medium_icon_button": 36,
        }
        self._popup_gap = self.metrics["popup_gap"]
        self._popup_large_gap = self.metrics["popup_large_gap"]

        self.left_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        self.center_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        self.right_wing = ctk.CTkFrame(self.pill_frame, fg_color="transparent")

        self.pill_frame.configure(
            fg_color=self.theme["bar_bg"],
            border_color=self.theme["bar_border"],
            corner_radius=self.metrics["bar_radius"],
        )
        self.bar_highlight = ctk.CTkFrame(
            self.pill_frame,
            height=self.metrics["bar_highlight_h"],
            fg_color=self.theme["bar_highlight"],
            corner_radius=0,
        )
        self.bar_highlight.place(
            relx=0.5,
            y=1,
            relwidth=0.92,
            anchor="n",
        )
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = int(sw * (float(self.config.get("width_percent", 90)) / 100))
        h = int(self.config.get("height", 50))

        if self.config.get("position", "Top") == "Bottom":
            y_pos = sh - h - self._edge_gap
        else:
            y_pos = 0

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

        # Design tokens derived from the configured bar height.
        _bg = self.theme["bar_bg"]
        _bh = max(26, h - 12)  # button height: 33px at default h=45
        _icr = self.metrics["button_radius"]
        _icf = self.fonts["icon_sm"]
        _ich = self.theme["text_dim"]  # icon glyph colour
        _ihv = self.theme["surface_hover"]  # hover background
        self._btn_h = _bh  # stored so add_app_button can use it
        self._app_btn_w = (
            _bh + 8
        )  # app slots need a little extra width so icons don't clip
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

        # Widgets
        # Active window title
        self.active_window_label = ctk.CTkLabel(
            self.pill_frame,
            text="",
            height=_bh,
            font=self.fonts["bar_label"],
            text_color=self.theme["text_muted"],
            anchor="w",
        )

        # Clock - Canvas for pixel-exact two-line vertical centering
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
            font=self.fonts["bar_label_bold"],
            anchor="center",
        )
        self._clock_date_id = self.clock_container.create_text(
            _cx,
            _cy + 8,
            text="",
            fill=self.theme["text_faint"],
            font=self.fonts["bar_minor"],
            anchor="center",
        )

        # System monitor - Canvas, two-line, mirrors clock layout
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
            font=self.fonts["bar_meter"],
            anchor="center",
        )
        self._sys_ram_id = self.sys_container.create_text(
            _sx,
            _sy + 8,
            text="RAM  0%",
            fill=self.theme["text_faint"],
            font=self.fonts["bar_meter"],
            anchor="center",
        )

        # Separator between sys_monitor and clock - inserted by render_bar
        self._info_sep = tk.Frame(
            self.pill_frame,
            width=self.metrics["line_w"],
            height=self.metrics["separator_h"],
            bg=self.theme["separator"],
            bd=0,
            highlightthickness=0,
        )

        # App taskbar
        self.apps_container = ctk.CTkFrame(self.pill_frame, fg_color="transparent")
        self.overflow_btn = _ibtn(
            self.apps_container, "\ue712", self.toggle_overflow_menu
        )
        self.overflow_apps = []

        # Icon buttons - all identical sizing/style
        self.start_btn = _ibtn(
            self.pill_frame,
            "\ue782",
            self.toggle_custom_start,
            font=self.fonts["icon_bar"],
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

        # Edit-mode "Done" button - same height as icon buttons
        self.edit_done_btn = ctk.CTkButton(
            self.pill_frame,
            text="Done",
            width=self.metrics["done_button_w"],
            height=_bh,
            fg_color=self.theme["accent"],
            hover_color=self.theme["accent_hover"],
            corner_radius=_icr,
            font=self.fonts["caption_bold"],
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
        self._open_apps_scan_inflight = False
        self._open_apps_signature = None
        self._process_exe_cache = {}
        self._last_real_title: str = ""  # last focused title from a non-WinBar window
        self._active_hwnd = None  # hwnd of the currently highlighted app button
        self.app_groups = {}
        self.hwnd_to_group = {}
        self._group_icon_cache = {}
        self._startup_icon_verify_scheduled = False
        self._startup_icon_verify_done = False
        self._hovered_group_key = None
        self._hovered_group_anchor = None
        self._stack_menu_anchor = None
        self._stack_menu_group = None
        self._stack_menu_after = None
        self._stack_menu_hide_after = None
        self._pinned_buttons = []
        self._pinned_separator = None
        self._resolved_icon_paths = {}
        self._exe_icon_cache = {}
        self._ctk_icon_cache = {}
        self._icon_load_inflight = set()
        self._group_icon_retry_counts = {}
        self._volume_available = _AUDIO_API_AVAILABLE
        self._volume_slider_busy = False
        self._volume_poll_inflight = False
        self._volume_set_inflight = False
        self._mute_set_inflight = False
        self._pending_volume_scalar = None
        self._pending_mute_value = None
        self._volume_state_cache = (0.0, False)
        self._volume_device_cache = "Default output"
        self._background_icon_cache = {}
        self._background_icon_load_inflight = set()
        self._background_apps_cache = None
        self._background_apps_cache_signature = None
        self._background_apps_cache_time = 0.0
        self._background_apps_cache_ttl = 8.0
        self._background_apps_refresh_inflight = False
        self._background_apps_request = 0
        self._search_request = 0
        self._search_after_id = None
        self._search_last_render_key = None
        self._last_search_query = None
        self._start_pins_request = 0
        self._start_menu_open_request = 0
        self._tooltip = None
        self._tooltip_after = None
        self._tooltip_hide_after = None
        self._tooltip_anchor = None
        self._tooltip_text = None
        self._ui_queue = queue.Queue()
        self._ui_queue_after_id = None
        self._ui_queue_interval_ms = 33
        self._taskbar_guard_after_id = None
        self._taskbar_guard_interval_ms = 1000
        self._taskbar_guard_ticks = 0
        self._last_clock_text = None
        self._last_date_text = None
        self._last_cpu_text = None
        self._last_ram_text = None
        self._last_active_label_text = None
        self._last_volume_button_refresh = 0.0
        self._volume_button_refresh_interval = 5.0
        self._open_apps_interval_ms = 1000
        self._foreground_poll_ms = 250
        self._fullscreen_poll_ms = 1000

        psutil.cpu_percent(interval=None)
        # Pre-build the app search index off the main thread so first search is instant
        _start_app_index_build()
        self.render_bar()
        self._schedule_ui_queue_drain()
        self._schedule_taskbar_guard(150)
        self._schedule_volume_button_refresh()
        self.after(1200, self._prewarm_popup_resources)
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
        if popup is None or not popup.winfo_exists():
            return False
        try:
            return popup.state() != "withdrawn"
        except Exception:
            return True

    def _destroy_popup(self, attr):
        popup = getattr(self, attr, None)
        if popup is not None:
            try:
                if popup.winfo_exists():
                    if attr == "search_window" and not getattr(self, "_exiting", False):
                        after_id = getattr(self, "_search_after_id", None)
                        if after_id is not None:
                            try:
                                self.after_cancel(after_id)
                            except Exception:
                                pass
                            self._search_after_id = None
                        popup.wm_attributes("-alpha", 0.0)
                        popup.withdraw()
                        return
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

    def _is_pointer_over_widget(self, widget, pad=0):
        if widget is None or not widget.winfo_exists():
            return False
        mx, my = self.winfo_pointerxy()
        wx = widget.winfo_rootx()
        wy = widget.winfo_rooty()
        return (
            wx - pad <= mx <= wx + widget.winfo_width() + pad
            and wy - pad <= my <= wy + widget.winfo_height() + pad
        )

    def _popup_anchor_y(self, anchor_y, anchor_h, menu_h, gap=None):
        gap = self._popup_gap if gap is None else gap
        if self.config.get("position", "Top") == "Bottom":
            return anchor_y - menu_h - gap
        return anchor_y + anchor_h + gap

    def _popup_geometry(
        self,
        menu_w,
        menu_h,
        *,
        anchor_widget=None,
        align="left",
        gap=None,
        offset_x=0,
        x=None,
        y=None,
    ):
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
        return f"{menu_w}x{menu_h}+{x}+{y}"

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
        corner_radius=None,
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
            0.0,
        )

        popup.geometry(
            self._popup_geometry(
                menu_w,
                menu_h,
                anchor_widget=anchor_widget,
                align=align,
                gap=gap,
                offset_x=offset_x,
                x=x,
                y=y,
            )
        )

        panel = ctk.CTkFrame(
            popup,
            corner_radius=corner_radius or self.metrics["popup_radius"],
            fg_color=self.theme["surface"],
            border_width=1,
            border_color=self.theme["bar_border"],
        )
        panel.pack(fill="both", expand=True)
        try:
            popup.deiconify()
            popup.lift()
        except Exception:
            pass
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
            if anchor_widget is not None and self._is_pointer_over_widget(
                anchor_widget
            ):
                return
            self._destroy_popup(popup_attr)

        popup.bind("<FocusOut>", lambda _event: popup.after(35, _close_if_outside))

    def _popup_header(self, parent, icon, title, subtitle=None, value=None):
        header = ctk.CTkFrame(parent, fg_color="transparent")
        header.pack(
            fill="x",
            padx=self.metrics["popup_header_pad_x"],
            pady=self.metrics["popup_header_pad_y"],
        )

        icon_chip = ctk.CTkFrame(
            header,
            width=self.metrics["chip_size"],
            height=self.metrics["chip_size"],
            corner_radius=self.metrics["chip_radius"],
            fg_color=self.theme["surface_alt"],
            border_width=1,
            border_color=self.theme["separator"],
        )
        icon_chip.pack(side="left")
        icon_chip.pack_propagate(False)
        icon_label = ctk.CTkLabel(
            icon_chip,
            text=icon,
            font=self.fonts["icon_sm"],
            text_color=self.theme["text_dim"],
        )
        icon_label.pack(expand=True)

        title_wrap = ctk.CTkFrame(header, fg_color="transparent")
        title_wrap.pack(side="left", fill="x", expand=True, padx=(10, 0))
        ctk.CTkLabel(
            title_wrap,
            text=title,
            font=self.fonts["title"],
            text_color=self.theme["text"],
            anchor="w",
        ).pack(fill="x")
        if subtitle:
            ctk.CTkLabel(
                title_wrap,
                text=subtitle,
                font=self.fonts["caption"],
                text_color=self.theme["text_muted"],
                anchor="w",
            ).pack(fill="x")

        if value is not None:
            value_label = ctk.CTkLabel(
                header,
                text=value,
                font=self.fonts["label_bold"],
                text_color=self.theme["text_muted"],
            )
            value_label.pack(side="right")
            return icon_label, value_label
        return icon_label, None

    def _popup_divider(self, parent, pady=(0, 10)):
        div = tk.Frame(
            parent,
            height=1,
            bg=self.theme["separator"],
            bd=0,
            highlightthickness=0,
        )
        div.pack(fill="x", padx=self.metrics["divider_pad_x"], pady=pady)

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
                "text_color": self.theme["danger_text"],
            },
        }[variant]
        btn_font = kwargs.pop("font", self.fonts["caption_bold"])
        options = {**base, **kwargs}
        return ctk.CTkButton(
            parent,
            corner_radius=self.metrics["row_radius"],
            border_width=0,
            font=btn_font,
            **options,
        )

    def _post_ui(self, callback):
        if getattr(self, "_exiting", False):
            return
        try:
            self._ui_queue.put_nowait(callback)
        except Exception:
            pass

    def _schedule_ui_queue_drain(self):
        if getattr(self, "_exiting", False):
            return
        try:
            self._ui_queue_after_id = self.after(
                self._ui_queue_interval_ms, self._drain_ui_queue
            )
        except Exception:
            self._ui_queue_after_id = None

    def _drain_ui_queue(self):
        if getattr(self, "_exiting", False):
            return

        for _ in range(80):
            try:
                callback = self._ui_queue.get_nowait()
            except queue.Empty:
                break

            try:
                if self.winfo_exists():
                    callback()
            except Exception as exc:
                log_error("_drain_ui_queue", exc)

        self._schedule_ui_queue_drain()

    def _resolve_icon_path(self, path):
        if not path:
            return None
        key = os.path.normcase(os.path.abspath(path))
        if key in self._resolved_icon_paths:
            return self._resolved_icon_paths[key] or None

        resolved = resolve_lnk(path) if path.lower().endswith(".lnk") else path
        if not resolved or not os.path.exists(resolved):
            resolved = ""
        self._resolved_icon_paths[key] = resolved
        return resolved or None

    def _get_cached_ctk_icon(self, path, size, *, load=False):
        resolved = self._resolve_icon_path(path)
        if not resolved:
            return None

        norm = os.path.normcase(resolved)
        cached_img = self._exe_icon_cache.get(norm)
        if cached_img is None and load:
            cached_img = get_icon_from_exe(resolved) or False
            self._exe_icon_cache[norm] = cached_img

        if not cached_img:
            return None

        icon_key = (norm, int(size))
        icon = self._ctk_icon_cache.get(icon_key)
        if icon is None:
            icon = ctk.CTkImage(cached_img, cached_img, size=(int(size), int(size)))
            self._ctk_icon_cache[icon_key] = icon
        return icon

    def _get_ready_ctk_icon(self, path, size):
        if not path:
            return None
        key = os.path.normcase(os.path.abspath(path))
        if key not in self._resolved_icon_paths:
            return None
        return self._get_cached_ctk_icon(path, size, load=False)

    def _get_ctk_icon_from_image(self, key, img, size):
        if not img:
            return None
        icon_key = (key, int(size), "ctk")
        icon = self._ctk_icon_cache.get(icon_key)
        if icon is None:
            try:
                icon = ctk.CTkImage(img, img, size=(int(size), int(size)))
                self._ctk_icon_cache[icon_key] = icon
            except Exception as exc:
                log_error("_get_ctk_icon_from_image", exc)
                return None
        return icon

    def _set_button_icon(self, button, icon, text=""):
        if button is None or not button.winfo_exists() or icon is None:
            return
        button.configure(require_redraw=True, image=icon, text=text)
        button._icon_ref = icon
        self._refresh_button_surface_bindings(button)
        try:
            button.update_idletasks()
        except Exception:
            pass

    def _bind_button_surface(self, button, bindings):
        button._winbar_surface_bindings = tuple(bindings)
        button._winbar_bound_image_label = getattr(button, "_image_label", None)
        button._winbar_bound_text_label = getattr(button, "_text_label", None)
        for sequence, callback in button._winbar_surface_bindings:
            button.bind(sequence, callback)

    def _refresh_button_surface_bindings(self, button):
        bindings = getattr(button, "_winbar_surface_bindings", ())
        if not bindings:
            return
        for attr, marker in (
            ("_image_label", "_winbar_bound_image_label"),
            ("_text_label", "_winbar_bound_text_label"),
        ):
            label = getattr(button, attr, None)
            if label is None or getattr(button, marker, None) is label:
                continue
            for sequence, callback in bindings:
                label.bind(sequence, callback, add=True)
            setattr(button, marker, label)

    def _apply_button_icon(
        self, button, path, size, popup_attr, request_attr, request_id, text_with_icon
    ):
        if not button.winfo_exists():
            return
        if popup_attr and not self._popup_exists(popup_attr):
            return
        if request_attr and getattr(self, request_attr, None) != request_id:
            return

        icon = self._get_cached_ctk_icon(path, size, load=False)
        if icon is None:
            return
        self._set_button_icon(button, icon, text_with_icon or "")
        if text_with_icon is not None:
            button.configure(text=text_with_icon)

    def _apply_ready_pinned_icons(self):
        for container in getattr(self, "_pinned_buttons", []):
            try:
                if not container.winfo_exists():
                    continue
                path = getattr(container, "_icon_source_path", None)
                size = getattr(container, "_icon_size", None)
                if not path or not size:
                    continue
                resolved = self._resolve_icon_path(path)
                if not resolved:
                    continue
                norm = os.path.normcase(resolved)
                img = self._exe_icon_cache.get(norm)
                if not img:
                    continue
                button = getattr(container, "_button", None)
                if button is not None and button.winfo_exists():
                    icon = self._get_cached_ctk_icon(path, size, load=False)
                    if icon is None:
                        continue
                    self._set_button_icon(button, icon)
                    try:
                        container.update_idletasks()
                    except Exception:
                        pass
            except Exception as exc:
                log_error("_apply_ready_pinned_icons", exc)

    def _cancel_pinned_button_animation(self, container):
        after_ids = getattr(container, "_pinned_anim_after_ids", ())
        for after_id in after_ids:
            try:
                self.after_cancel(after_id)
            except Exception:
                pass
        container._pinned_anim_after_ids = []

    def _animate_pinned_indicator(self, container, show):
        indicator = getattr(container, "_indicator", None)
        if indicator is None or not indicator.winfo_exists():
            return

        self._cancel_pinned_button_animation(container)

        bh = self._btn_h
        visible_y = bh - 4
        hidden_y = bh + 2
        start_y = int(getattr(container, "_indicator_y", hidden_y))
        target_y = visible_y if show else hidden_y

        if show:
            try:
                indicator.place(relx=0.5, y=start_y, anchor="n")
            except Exception:
                pass

        steps = max(1, int(self.metrics["pinned_anim_steps"]))
        step_ms = max(1, int(self.metrics["pinned_anim_step_ms"]))
        after_ids = []

        def _apply(progress):
            eased = 1 - ((1 - progress) * (1 - progress))
            y = int(round(start_y + (target_y - start_y) * eased))
            try:
                indicator.place_configure(y=y)
                container._indicator_y = y
            except Exception:
                pass
            if progress >= 1 and not show:
                try:
                    indicator.place_forget()
                except Exception:
                    pass

        for step in range(1, steps + 1):

            def _step(progress=step / steps):
                try:
                    if container.winfo_exists():
                        _apply(progress)
                except Exception as exc:
                    log_error("_animate_pinned_indicator", exc)
                finally:
                    if progress >= 1:
                        container._pinned_anim_after_ids = []

            after_ids.append(self.after(step * step_ms, _step))

        container._pinned_anim_after_ids = after_ids

    def _set_pinned_button_state(self, container, state="inactive"):
        button = getattr(container, "_button", None)
        indicator = getattr(container, "_indicator", None)
        if button is None or not button.winfo_exists():
            return
        if getattr(container, "_pinned_state", None) == state:
            return
        container._pinned_state = state

        if state == "pressed":
            button.configure(
                fg_color=self.theme["pinned_press"],
                hover_color=self.theme["pinned_press"],
                border_color=self.theme["surface_active_border"],
            )
            if indicator is not None and indicator.winfo_exists():
                indicator.configure(
                    fg_color=self.theme["surface_active_border"],
                    width=self.metrics["indicator_active_w"],
                )
            self._animate_pinned_indicator(container, show=True)
        elif state == "hover":
            button.configure(
                fg_color=self.theme["pinned_hover"],
                hover_color=self.theme["pinned_hover"],
                border_color=self.theme["surface_active_border"],
            )
            if indicator is not None and indicator.winfo_exists():
                indicator.configure(
                    fg_color=self.theme["surface_active_border"],
                    width=self.metrics["indicator_group_w"],
                )
            self._animate_pinned_indicator(container, show=True)
        else:
            button.configure(
                fg_color=self.theme["surface_alt"],
                hover_color=self.theme["surface_alt"],
                border_color=self.theme["separator"],
            )
            self._animate_pinned_indicator(container, show=False)

    def _leave_pinned_launcher(self, container):
        if container is None or not container.winfo_exists():
            return
        try:
            if self._is_pointer_over_widget(container):
                self._set_pinned_button_state(container, "hover")
                return
        except Exception:
            pass
        self._set_pinned_button_state(container, "inactive")
        self._hide_tooltip()

    def _load_button_icon_async(
        self,
        button,
        path,
        size,
        *,
        popup_attr=None,
        request_attr=None,
        request_id=None,
        text_with_icon=None,
    ):
        if not button.winfo_exists():
            return
        if popup_attr and not self._popup_exists(popup_attr):
            return
        if request_attr and getattr(self, request_attr, None) != request_id:
            return

        icon = self._get_cached_ctk_icon(path, size, load=False)
        if icon is not None:
            self._apply_button_icon(
                button, path, size, popup_attr, request_attr, request_id, text_with_icon
            )
            return

        resolved = self._resolve_icon_path(path)
        if not resolved:
            return

        norm = os.path.normcase(resolved)
        if norm in self._icon_load_inflight:
            self.after(
                35,
                lambda: self._load_button_icon_async(
                    button,
                    path,
                    size,
                    popup_attr=popup_attr,
                    request_attr=request_attr,
                    request_id=request_id,
                    text_with_icon=text_with_icon,
                ),
            )
            return

        self._icon_load_inflight.add(norm)

        def _worker():
            img = get_icon_from_exe(resolved)

            def _apply():
                self._icon_load_inflight.discard(norm)
                self._exe_icon_cache[norm] = img or False
                self._apply_button_icon(
                    button,
                    path,
                    size,
                    popup_attr,
                    request_attr,
                    request_id,
                    text_with_icon,
                )
                self._apply_ready_pinned_icons()

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _prime_icon_async(self, path):
        resolved = self._resolve_icon_path(path)
        if not resolved:
            return
        norm = os.path.normcase(resolved)
        if norm in self._exe_icon_cache or norm in self._icon_load_inflight:
            return

        self._icon_load_inflight.add(norm)

        def _worker():
            img = get_icon_from_exe(resolved)

            def _apply():
                self._icon_load_inflight.discard(norm)
                self._exe_icon_cache[norm] = img or False
                self._apply_ready_pinned_icons()

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _prewarm_popup_resources(self):
        for app in self.config.get("pinned_apps", []):
            path = app.get("path")
            if path:
                self._prime_icon_async(path)
        self._load_background_apps_async()

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

    def _schedule_taskbar_guard(self, delay_ms=None):
        if getattr(self, "_exiting", False):
            return
        delay = self._taskbar_guard_interval_ms if delay_ms is None else delay_ms
        try:
            self._taskbar_guard_after_id = self.after(delay, self._taskbar_guard)
        except Exception:
            self._taskbar_guard_after_id = None

    def _taskbar_guard(self):
        self._taskbar_guard_after_id = None
        if getattr(self, "_exiting", False):
            return

        try:
            self._taskbar_guard_ticks += 1
            should_reassert = (
                is_taskbar_visible() or self._taskbar_guard_ticks % 10 == 0
            )
            if should_reassert:
                set_taskbar_visibility(False)

            if self._my_hwnd is None:
                self._my_hwnd = ctypes.windll.user32.GetAncestor(
                    self.winfo_id(), GA_ROOT
                )
            if should_reassert and float(self.attributes("-alpha") or 1.0) > 0.05:
                HWND_TOPMOST = -1
                FLAGS = 0x0002 | 0x0001 | 0x0010
                ctypes.windll.user32.SetWindowPos(
                    self._my_hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
                )
        except Exception as exc:
            log_error("_taskbar_guard", exc)

        self._schedule_taskbar_guard()

    def update_time(self):
        now = datetime.datetime.now()
        cpu = int(psutil.cpu_percent())
        ram = int(psutil.virtual_memory().percent)
        clock_text = now.strftime("%#I:%M %p")
        date_text = now.strftime("%a, %b %#d")
        cpu_text = f"CPU {cpu:3}%"
        ram_text = f"RAM {ram:3}%"
        if clock_text != self._last_clock_text:
            self.clock_container.itemconfig(self._clock_time_id, text=clock_text)
            self._last_clock_text = clock_text
        if date_text != self._last_date_text:
            self.clock_container.itemconfig(self._clock_date_id, text=date_text)
            self._last_date_text = date_text
        if cpu_text != self._last_cpu_text:
            self.sys_container.itemconfig(self._sys_cpu_id, text=cpu_text)
            self._last_cpu_text = cpu_text
        if ram_text != self._last_ram_text:
            self.sys_container.itemconfig(self._sys_ram_id, text=ram_text)
            self._last_ram_text = ram_text
        self._sync_volume_button_icon()
        now_monotonic = time.monotonic()
        if (
            now_monotonic - self._last_volume_button_refresh
            >= self._volume_button_refresh_interval
        ):
            self._last_volume_button_refresh = now_monotonic
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
                    label_text = f"\ue76c  {trimmed}"
                else:
                    label_text = ""
                if label_text != self._last_active_label_text:
                    self.active_window_label.configure(text=label_text)
                    self._last_active_label_text = label_text
            elif not hwnd:
                self._update_active_indicator(0)
                if self._last_active_label_text != "":
                    self.active_window_label.configure(text="")
                    self._last_active_label_text = ""
        except Exception:
            pass
        self.after(self._foreground_poll_ms, self._poll_foreground)

    def _update_active_indicator(self, hwnd: int):
        if hwnd == self._active_hwnd:
            return
        # Clear old indicator - restore to the running-app resting state
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

    def _get_exe_from_hwnd_cached(self, hwnd, now=None):
        now = time.monotonic() if now is None else now
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
        except Exception:
            return None
        if not pid:
            return None

        cached = self._process_exe_cache.get(pid)
        if cached and now - cached[1] < 12.0:
            return cached[0] or None

        try:
            exe = psutil.Process(pid).exe()
        except Exception:
            exe = None
        self._process_exe_cache[pid] = (exe, now)

        if len(self._process_exe_cache) > 256:
            cutoff = now - 30.0
            self._process_exe_cache = {
                key: value
                for key, value in self._process_exe_cache.items()
                if value[1] >= cutoff
            }
        return exe

    def _build_open_app_groups(self, current_windows):
        _skip = {
            "FloatingBar",
            "Limitens",
            "NVIDIA",
            "Overlay",
            "CTkToplevel",
            "Windows Shell Experience Host",
        }
        grouped = {}
        now = time.monotonic()
        for app in current_windows:
            if any(x in app["title"] for x in _skip):
                continue
            hid = app["hwnd"]
            exe = self._get_exe_from_hwnd_cached(hid, now)
            group_key = (
                os.path.normcase(exe) if exe else f"title:{app['title'].lower()}"
            )
            group_name = (
                os.path.splitext(os.path.basename(exe))[0] if exe else app["title"][:20]
            )
            group = grouped.setdefault(
                group_key,
                {"name": group_name, "exe": exe, "windows": []},
            )
            group["windows"].append({"hwnd": hid, "title": app["title"]})
        return grouped

    def _open_app_signature(self, grouped):
        return tuple(
            sorted(
                (
                    group_key,
                    tuple(sorted(item["hwnd"] for item in group["windows"])),
                    group.get("exe") or "",
                )
                for group_key, group in grouped.items()
            )
        )

    def _start_group_icon_load(self, group_key, exe=None, hwnds=None, force=False):
        hwnds = tuple(hwnds or ())
        if group_key in self._pending_icons or (not exe and not hwnds):
            return
        if force:
            self._exe_icon_cache.pop(group_key, None)
            self._group_icon_retry_counts.pop(group_key, None)
        self._pending_icons.add(group_key)

        def _load(g=group_key, e=exe, window_handles=hwnds):
            img = get_icon_from_exe(e) if e else None
            if img is None:
                for hwnd in window_handles:
                    img = get_icon_from_hwnd(hwnd)
                    if img is not None:
                        break

            def _apply():
                if img:
                    self._exe_icon_cache[g] = img
                    self._group_icon_retry_counts.pop(g, None)
                    self._sync_group_icon(g, img)
                else:
                    self._pending_icons.discard(g)
                    retries = self._group_icon_retry_counts.get(g, 0) + 1
                    if retries >= 5:
                        self._exe_icon_cache[g] = False
                        self._group_icon_retry_counts.pop(g, None)
                    else:
                        self._group_icon_retry_counts[g] = retries
                        self.after(
                            400,
                            lambda gg=g, ee=e, hs=window_handles: (
                                self._retry_group_icon_load(gg, ee, hs)
                            ),
                        )

            self._post_ui(_apply)

        threading.Thread(target=_load, daemon=True).start()

    def _retry_group_icon_load(self, group_key, exe=None, hwnds=None):
        if group_key in self._group_icon_cache or group_key not in self.app_groups:
            return
        if self._exe_icon_cache.get(group_key) is False:
            return
        self._start_group_icon_load(group_key, exe, hwnds)

    def _register_group_icons(self, grouped, *, force=False):
        requested = 0
        for group_key, group in grouped.items():
            exe = group.get("exe")
            hwnds = [item["hwnd"] for item in group.get("windows", [])]
            if group_key in self._group_icon_cache and not force:
                continue
            cached_icon = self._exe_icon_cache.get(group_key)
            if cached_icon and not force:
                self._sync_group_icon(group_key, cached_icon)
                continue
            if cached_icon is False and not force:
                continue
            if exe or hwnds:
                requested += 1
                self._start_group_icon_load(group_key, exe, hwnds, force=force)
        return requested

    def _schedule_startup_icon_verification(self):
        if self._startup_icon_verify_scheduled or self._startup_icon_verify_done:
            return
        self._startup_icon_verify_scheduled = True
        self.after(900, lambda: self._verify_registered_icons(attempt=1))

    def _verify_registered_icons(self, attempt=1):
        if getattr(self, "_exiting", False):
            return
        total = 0
        missing = []
        for group_key, container in getattr(self, "active_app_buttons", {}).items():
            if not container.winfo_exists():
                continue
            button = getattr(container, "_button", None)
            if button is None or not button.winfo_exists():
                continue
            total += 1
            if not button.cget("image"):
                missing.append(group_key)

        if missing and attempt <= 4:
            groups = {
                group_key: self.app_groups[group_key]
                for group_key in missing
                if group_key in self.app_groups
            }
            self._register_group_icons(groups, force=True)
            self.after(700, lambda a=attempt + 1: self._verify_registered_icons(a))
            return

        loaded = max(0, total - len(missing))
        self._startup_icon_verify_done = True
        log_event(
            "startup_icon_verify",
            f"loaded={loaded} missing={len(missing)} total={total} attempts={attempt}",
        )

    def _sync_group_icon(self, group_key, img):
        self._pending_icons.discard(group_key)
        if not img:
            return
        self._group_icon_cache[group_key] = img
        if group_key and not str(group_key).startswith("title:"):
            self._exe_icon_cache[group_key] = img
        container = self.active_app_buttons.get(group_key)
        button = getattr(container, "_button", None)
        if button is not None and button.winfo_exists():
            icon_size = int(self._btn_h * 0.55)
            icon = self._get_ctk_icon_from_image(group_key, img, icon_size)
            if icon is not None:
                self._set_button_icon(button, icon)
                container._icon_ref = icon
            try:
                container.update_idletasks()
            except Exception:
                pass
            return
        self._rebuild_group_buttons()

    def _handle_group_enter(self, widget, group_key):
        group = self.app_groups.get(group_key)
        if not group:
            return
        self._hovered_group_key = group_key
        self._hovered_group_anchor = widget
        self._cancel_tooltip_hide()
        self._cancel_stack_menu_hide()
        if len(group.get("windows", [])) > 1:
            self._hide_tooltip()
            self._schedule_stack_menu(widget, group_key)
        else:
            self._hide_stack_menu(clear_hover=False)
            self._schedule_tooltip(widget, self._group_label(group))

    def _handle_group_leave(self, widget=None):
        def _clear_if_outside(anchor=widget, group_key=self._hovered_group_key):
            stack_menu = getattr(self, "_stack_menu", None)
            inside_anchor = anchor is not None and self._is_pointer_over_widget(
                anchor, pad=8
            )
            inside_stack = stack_menu is not None and self._is_pointer_over_widget(
                stack_menu, pad=8
            )
            if inside_anchor or inside_stack:
                return
            if self._hovered_group_key == group_key:
                self._hovered_group_key = None
                self._hovered_group_anchor = None
            self._schedule_tooltip_hide(anchor)
            self._schedule_stack_menu_hide()

        self.after(60, _clear_if_outside)

    def _set_group_button_state(self, container, active=False):
        button = getattr(container, "_button", None)
        indicator = getattr(container, "_indicator", None)
        count = getattr(container, "_window_count", 1)
        bh = self._btn_h
        if button and button.winfo_exists():
            if active:
                button.configure(
                    fg_color=self.theme["surface_active"],
                    hover_color=self.theme["surface_active"],
                    border_width=1,
                    border_color=self.theme["surface_active_border"],
                )
            else:
                button.configure(
                    fg_color=self.theme["surface_alt"],
                    hover_color=self.theme["surface_hover"],
                    border_width=1,
                    border_color=self.theme["separator"],
                )
        if indicator and indicator.winfo_exists():
            if active:
                indicator.configure(
                    fg_color=self.theme["surface_active_border"],
                    width=self.metrics["indicator_active_w"],
                )
                indicator.place(relx=0.5, y=bh - 4, anchor="n")
            elif count > 1:
                indicator.configure(
                    fg_color=self.theme["text_muted"],
                    width=self.metrics["indicator_group_w"],
                )
                indicator.place(relx=0.5, y=bh - 4, anchor="n")
            else:
                indicator.place_forget()

    def _create_group_button(self, group_key, group):
        bh = self._btn_h
        bw = self._app_btn_w
        icon_size = int(bh * 0.55)
        img = self._group_icon_cache.get(group_key)
        icon = None
        if img:
            icon = self._get_ctk_icon_from_image(group_key, img, icon_size)
        elif group.get("exe"):
            cached_img = self._exe_icon_cache.get(os.path.normcase(group["exe"]))
            if cached_img:
                icon = self._get_ctk_icon_from_image(group_key, cached_img, icon_size)
        count = len(group.get("windows", []))
        container = ctk.CTkFrame(
            self.apps_container, fg_color="transparent", width=bw + 4, height=bh
        )
        container.pack_propagate(False)
        btn = ctk.CTkButton(
            container,
            text="" if icon else "\ue7c3",
            image=icon,
            width=bw,
            height=bh,
            fg_color=self.theme["surface_alt"],
            hover_color=self.theme["surface_hover"],
            corner_radius=self.metrics["button_radius"],
            font=self.fonts["icon_sm"],
            text_color=self.theme["text_dim"],
            border_width=1,
            border_color=self.theme["separator"],
            anchor="center",
            border_spacing=0,
        )
        btn.place(relx=0.5, y=0, anchor="n")
        if icon:
            btn._icon_ref = icon
        indicator = ctk.CTkFrame(
            container,
            width=self.metrics["indicator_w"],
            height=self.metrics["indicator_h"],
            corner_radius=self.metrics["indicator_radius"],
            fg_color=self.theme["text_muted"],
        )

        close_group = lambda e, g=group_key: self._close_group_window(g)
        focus_group = lambda e, g=group_key: self.focus_group(g)
        enter_group = lambda e, w=container, g=group_key: self._handle_group_enter(w, g)
        leave_group = lambda e, w=container: self._handle_group_leave(w)

        for widget in (container, indicator):
            widget.bind("<Button-2>", close_group)
            widget.bind("<Button-1>", focus_group)
            widget.bind("<Enter>", enter_group)
            widget.bind("<Leave>", leave_group)
        self._bind_button_surface(
            btn,
            (
                ("<Button-2>", close_group),
                ("<Button-1>", focus_group),
                ("<Enter>", enter_group),
                ("<Leave>", leave_group),
            ),
        )

        container._full_name = self._group_label(group)
        container._is_pinned = False
        container._group_key = group_key
        container._button = btn
        container._indicator = indicator
        container._window_count = count
        if icon:
            container._icon_ref = icon
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
        name = app.get("name", "App")
        icon = None
        resolved = self._resolve_icon_path(path)
        if resolved:
            cached_img = self._exe_icon_cache.get(os.path.normcase(resolved))
            if cached_img:
                icon = self._get_ctk_icon_from_image(
                    os.path.normcase(resolved), cached_img, icon_size
                )
        slot_width = bw + 2
        container = ctk.CTkFrame(
            self.apps_container,
            fg_color="transparent",
            width=slot_width,
            height=bh,
        )
        container.pack_propagate(False)
        btn = ctk.CTkButton(
            container,
            text="" if icon else name[:3],
            image=icon,
            width=bw,
            height=bh,
            fg_color=self.theme["surface_alt"],
            hover_color=self.theme["surface_alt"],
            corner_radius=self.metrics["button_radius"],
            font=self.fonts["caption"],
            text_color=self.theme["text_dim"],
            border_width=1,
            border_color=self.theme["separator"],
            anchor="center",
            border_spacing=0,
        )
        btn.place(relx=0.5, y=0, anchor="n")
        if icon:
            btn._icon_ref = icon
        indicator = ctk.CTkFrame(
            container,
            width=self.metrics["indicator_group_w"],
            height=self.metrics["indicator_h"],
            fg_color=self.theme["surface_active_border"],
            corner_radius=self.metrics["indicator_radius"],
        )

        def _launch(_event=None, p=path):
            if not self.edit_mode:
                self._launch_path(p)

        def _press(_event=None, frame=container):
            if not self.edit_mode:
                self._set_pinned_button_state(frame, "pressed")

        def _release(_event=None, frame=container):
            if self.edit_mode:
                return
            if self._is_pointer_over_widget(frame, pad=6):
                self._set_pinned_button_state(frame, "hover")
                _launch()
            else:
                self._set_pinned_button_state(frame, "inactive")

        enter_pinned = lambda e, w=container, t=name: [
            self._set_pinned_button_state(w, "hover"),
            self._schedule_tooltip(w, t),
        ]
        leave_pinned = lambda e, w=container: self.after(
            20, lambda b=w: self._leave_pinned_launcher(b)
        )

        for widget in (container, indicator):
            widget.bind("<ButtonPress-1>", _press)
            widget.bind("<ButtonRelease-1>", _release)
            widget.bind("<Enter>", enter_pinned)
            widget.bind("<Leave>", leave_pinned)
        self._bind_button_surface(
            btn,
            (
                ("<ButtonPress-1>", _press),
                ("<ButtonRelease-1>", _release),
                ("<Enter>", enter_pinned),
                ("<Leave>", leave_pinned),
            ),
        )

        container._full_name = name
        container._is_pinned = True
        container._icon_source_path = path
        container._icon_size = icon_size
        container._button = btn
        container._indicator = indicator
        container._indicator_y = self._btn_h + 2
        container._pinned_anim_after_ids = []
        if icon:
            container._icon_ref = icon
        if icon is None:
            self.after_idle(lambda p=path: self._prime_icon_async(p))
        self._set_pinned_button_state(container, "inactive")
        return container

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
                width=self.metrics["pinned_separator_w"],
                height=self._btn_h,
                fg_color="transparent",
            )
            self._pinned_separator.pack_propagate(False)
            ctk.CTkFrame(
                self._pinned_separator,
                width=1,
                height=max(10, self._btn_h // 2),
                fg_color=self.theme["separator"],
                corner_radius=0,
            ).place(relx=0.5, rely=0.5, anchor="center")

        self._repack_app_buttons()
        self._apply_ready_pinned_icons()

    def _repack_app_buttons(self):
        if not hasattr(self, "active_app_buttons"):
            return

        item_padx = 2
        pinned_padx = 1
        separator_padx = (2, 4)

        for child in self.apps_container.winfo_children():
            child.pack_forget()

        for btn in self._pinned_buttons:
            if btn.winfo_exists():
                btn.pack(side="left", padx=pinned_padx)

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
        if self._hovered_group_key != group_key and not self._popup_exists(
            "_stack_menu"
        ):
            return
        if self._stack_menu_group == group_key and self._popup_exists("_stack_menu"):
            self._stack_menu_anchor = widget
            self._cancel_stack_menu_hide()
            try:
                self._stack_menu.lift()
            except Exception:
                pass
            return

        self._hide_stack_menu(clear_hover=False)
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
            corner_radius=self.metrics["popup_radius_compact"],
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
                height=self.metrics["menu_row_h"],
                fg_color=(
                    self.theme["surface_alt"]
                    if item["hwnd"] == self._active_hwnd
                    else "transparent"
                ),
                hover_color=self.theme["surface_hover"],
                corner_radius=self.metrics["button_radius"],
                font=self.fonts["label_bold"],
                text_color=self.theme["text_dim"],
                command=lambda h=item["hwnd"]: [
                    self.focus_window(h),
                    self._hide_stack_menu(),
                ],
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
        if self._stack_menu_group == group_key and self._popup_exists("_stack_menu"):
            self._stack_menu_anchor = widget
            return
        self._stack_menu_after = self.after(
            self._stack_menu_open_delay_ms,
            lambda w=widget, g=group_key: (
                self._show_stack_menu(w, g)
                if self._hovered_group_key == g or self._popup_exists("_stack_menu")
                else None
            ),
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
                and anchor.winfo_rootx() - 8
                <= px
                <= anchor.winfo_rootx() + anchor.winfo_width() + 8
                and anchor.winfo_rooty() - 8
                <= py
                <= anchor.winfo_rooty() + anchor.winfo_height() + 8
            )
            inside_menu = (
                menu is not None
                and menu.winfo_exists()
                and menu.winfo_rootx() - 8
                <= px
                <= menu.winfo_rootx() + menu.winfo_width() + 8
                and menu.winfo_rooty() - 8
                <= py
                <= menu.winfo_rooty() + menu.winfo_height() + 8
            )
            if not inside_anchor and not inside_menu:
                self._hide_stack_menu()
        except Exception:
            self._hide_stack_menu()

    def _hide_stack_menu(self, clear_hover=True):
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
        if clear_hover:
            self._hovered_group_key = None
            self._hovered_group_anchor = None

    def _populate_start_pins(self):
        if not hasattr(self, "start_scroll") or not self.start_scroll.winfo_exists():
            return

        self._start_pins_request += 1
        request_id = self._start_pins_request

        for child in self.start_scroll.winfo_children():
            child.destroy()

        pinned_apps = self.config.get("pinned_apps", [])
        if not pinned_apps:
            ctk.CTkLabel(
                self.start_scroll,
                text="\ue71d",
                font=self.fonts["icon_lg"],
                text_color=self.theme["text_faint"],
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.start_scroll,
                text="No pinned apps yet",
                font=self.fonts["body_bold"],
                text_color=self.theme["text_dim"],
            ).pack()
            ctk.CTkLabel(
                self.start_scroll,
                text="Use Add App to pin an executable",
                font=self.fonts["caption"],
                text_color=self.theme["text_muted"],
            ).pack(pady=(2, 0))
            return

        for idx, app in enumerate(pinned_apps):
            app_path = app.get("path", "")
            app_name = app.get("name", "App")
            if not os.path.exists(app_path):
                continue

            row = ctk.CTkFrame(
                self.start_scroll,
                fg_color=self.theme["surface_alt"],
                corner_radius=self.metrics["row_radius"],
                border_width=1,
                border_color=self.theme["separator"],
            )
            row.pack(fill="x", pady=2, padx=5)

            icon = self._get_ready_ctk_icon(app_path, 32)
            resolved_icon_path = self._resolve_icon_path(app_path)
            icon_cache_key = (
                os.path.normcase(resolved_icon_path) if resolved_icon_path else None
            )

            launch_btn = self._popup_button(
                row,
                text=f"  {app_name}",
                image=icon,
                anchor="w",
                height=self.metrics["menu_large_row_h"],
                font=self.fonts["body_bold"],
                text_color=self.theme["text"],
                command=lambda p=app_path: [
                    self._launch_path(p),
                    self._destroy_popup("start_menu"),
                ],
            )
            launch_btn.pack(side="left", fill="x", expand=True, padx=(4, 0), pady=4)
            if icon is None and icon_cache_key not in self._exe_icon_cache:
                self.after_idle(
                    lambda b=launch_btn, p=app_path, rid=request_id, name=app_name: (
                        self._load_button_icon_async(
                            b,
                            p,
                            32,
                            popup_attr="start_menu",
                            request_attr="_start_pins_request",
                            request_id=rid,
                            text_with_icon=f"  {name}",
                        )
                    )
                )

            self._popup_button(
                row,
                text="\ue74d",
                width=self.metrics["medium_icon_button"],
                height=self.metrics["medium_icon_button"],
                font=self.fonts["icon_sm"],
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

        pinned_apps = self.config.setdefault("pinned_apps", [])
        if is_pinned_app_duplicate(pinned_apps, path):
            if hasattr(self, "start_menu") and self.start_menu.winfo_exists():
                self.start_menu.focus_force()
            return

        name = os.path.splitext(os.path.basename(path))[0]
        pinned_apps.append({"name": name, "path": path})
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
        if self._open_apps_scan_inflight:
            self.after(self._open_apps_interval_ms, self.update_open_apps)
            return
        self._open_apps_scan_inflight = True

        def _worker():
            try:
                grouped = self._build_open_app_groups(get_running_apps())
            except Exception as exc:
                log_error("update_open_apps", exc)
                grouped = None
            self._post_ui(lambda g=grouped: self._apply_open_apps_scan(g))

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_open_apps_scan(self, grouped):
        self._open_apps_scan_inflight = False
        if grouped is None:
            self.after(self._open_apps_interval_ms, self.update_open_apps)
            return

        signature = self._open_app_signature(grouped)
        if signature != self._open_apps_signature:
            self._open_apps_signature = signature
            self._prev_hwnds = {
                item["hwnd"]
                for group in grouped.values()
                for item in group.get("windows", [])
            }
            self.app_groups = grouped
            self.hwnd_to_group = {
                item["hwnd"]: group_key
                for group_key, group in grouped.items()
                for item in group["windows"]
            }
            self._rebuild_group_buttons()

        self._register_group_icons(grouped)
        self._schedule_startup_icon_verification()

        self.after(self._open_apps_interval_ms, self.update_open_apps)

    def check_fullscreen(self):
        try:
            hwnd = win32gui.GetForegroundWindow()
            if not hwnd:
                self.after(self._fullscreen_poll_ms, self.check_fullscreen)
                return

            # Cache my_hwnd - it never changes after the window is created
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
                self.after(self._fullscreen_poll_ms, self.check_fullscreen)
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
        self.after(self._fullscreen_poll_ms, self.check_fullscreen)

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
        window.update_idletasks()
        window.wm_attributes("-alpha", target_alpha)
        window.lift()

    def _show_tooltip(self, widget, text):
        if not self.config.get("hover_tooltips", True):
            self._hide_tooltip()
            return
        if widget is None or not widget.winfo_exists():
            return
        if (
            self._tooltip is not None
            and self._tooltip.winfo_exists()
            and self._tooltip_anchor is widget
            and self._tooltip_text == text
        ):
            return

        self._hide_tooltip()
        tip = ctk.CTkToplevel(self)
        tip.overrideredirect(True)
        tip.wm_attributes("-topmost", True)
        tip.configure(fg_color=self._transparent_key)
        tip.wm_attributes("-transparentcolor", self._transparent_key)

        panel = ctk.CTkFrame(
            tip,
            corner_radius=self.metrics["tooltip_radius"],
            fg_color=self.theme["tooltip_bg"],
            border_width=1,
            border_color=self.theme["tooltip_border"],
        )
        panel.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(panel, fg_color="transparent", corner_radius=0)
        inner.pack(
            fill="both",
            expand=True,
            padx=self.metrics["popup_inner_pad"],
            pady=self.metrics["control_gap"],
        )
        ctk.CTkLabel(
            inner,
            text=text,
            font=self.fonts["caption"],
            text_color=self.theme["text_dim"],
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
        self._tooltip_anchor = widget
        self._tooltip_text = text

    def _schedule_tooltip(self, widget, text):
        if not self.config.get("hover_tooltips", True):
            self._hide_tooltip()
            return
        self._cancel_tooltip_hide()
        self._cancel_tooltip()
        self._tooltip_after = self.after(
            self._tooltip_delay_ms, lambda: self._show_tooltip(widget, text)
        )

    def _cancel_tooltip(self):
        after_id = getattr(self, "_tooltip_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
            self._tooltip_after = None

    def _cancel_tooltip_hide(self):
        after_id = getattr(self, "_tooltip_hide_after", None)
        if after_id is not None:
            self.after_cancel(after_id)
            self._tooltip_hide_after = None

    def _schedule_tooltip_hide(self, anchor_widget=None):
        self._cancel_tooltip_hide()

        def _hide_if_outside():
            self._tooltip_hide_after = None
            anchor = anchor_widget or getattr(self, "_tooltip_anchor", None)
            if anchor is not None and self._is_pointer_over_widget(anchor, pad=6):
                return
            self._hide_tooltip()

        self._tooltip_hide_after = self.after(
            self._tooltip_hide_delay_ms, _hide_if_outside
        )

    def _hide_tooltip(self):
        self._cancel_tooltip()
        self._cancel_tooltip_hide()
        tip = getattr(self, "_tooltip", None)
        if tip is not None:
            try:
                tip.destroy()
            except Exception:
                pass
            self._tooltip = None
        self._tooltip_anchor = None
        self._tooltip_text = None

    def toggle_custom_start(self):
        if self.edit_mode:
            return
        if self._popup_exists("start_menu"):
            self._destroy_popup("start_menu")
            return
        self._start_menu_open_request += 1
        request_id = self._start_menu_open_request
        self._close_all_popups(exclude=("start_menu",))

        if self._prepare_start_menu_icons_for_open(request_id):
            return
        self._open_custom_start_menu(request_id)

    def _prepare_start_menu_icons_for_open(self, request_id):
        missing = []
        seen = set()
        for app in self.config.get("pinned_apps", []):
            path = app.get("path", "")
            if not path or not os.path.exists(path):
                continue
            resolved = self._resolve_icon_path(path)
            if not resolved:
                continue
            norm = os.path.normcase(resolved)
            if norm in seen or norm in self._exe_icon_cache:
                continue
            seen.add(norm)
            missing.append((norm, resolved))

        if not missing:
            return False

        for norm, _resolved in missing:
            self._icon_load_inflight.add(norm)

        def _worker(items=tuple(missing), rid=request_id):
            for norm, resolved in items:
                if self._exe_icon_cache.get(norm) is not None:
                    continue
                img = get_icon_from_exe(resolved)
                self._exe_icon_cache[norm] = img or False

            def _apply():
                for norm, _resolved in items:
                    self._icon_load_inflight.discard(norm)
                if (
                    rid == self._start_menu_open_request
                    and not self._popup_exists("start_menu")
                    and not getattr(self, "_exiting", False)
                ):
                    self._open_custom_start_menu(rid)

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()
        return True

    def _open_custom_start_menu(self, request_id=None):
        if request_id is not None and request_id != self._start_menu_open_request:
            return
        menu_w = 300
        menu_h = 400
        self.start_menu, main_frame = self._create_popup_shell(
            "start_menu",
            menu_w,
            menu_h,
            anchor_widget=self.start_btn,
            align="left",
            corner_radius=self.metrics["popup_radius"],
        )

        current_user = os.environ.get("USERNAME", "User")
        self._popup_header(
            main_frame,
            "\ue77b",
            current_user,
            subtitle="Pinned apps and system controls",
        )
        self._popup_divider(main_frame, pady=(0, 8))

        pins_header = ctk.CTkFrame(main_frame, fg_color="transparent")
        pins_header.pack(fill="x", padx=self.metrics["section_pad_x"], pady=(0, 6))
        ctk.CTkLabel(
            pins_header,
            text="Pinned apps",
            font=self.fonts["label"],
            text_color=self.theme["text_muted"],
        ).pack(side="left")
        self._popup_button(
            pins_header,
            text="+ Add App",
            height=self.metrics["menu_button_h"],
            width=self.metrics["action_button_w"],
            variant="soft",
            command=self.add_pinned_app,
        ).pack(side="right")

        self.start_scroll = ctk.CTkScrollableFrame(
            main_frame, fg_color="transparent", corner_radius=0
        )
        self.start_scroll.pack(
            fill="both",
            expand=True,
            padx=self.metrics["popup_inner_pad"],
            pady=5,
        )
        self._populate_start_pins()

        bottom_bar = ctk.CTkFrame(
            main_frame,
            fg_color=self.theme["surface_alt"],
            corner_radius=self.metrics["panel_radius"],
            border_width=1,
            border_color=self.theme["separator"],
        )
        bottom_bar.pack(
            fill="x",
            side="bottom",
            padx=self.metrics["popup_inner_pad"],
            pady=self.metrics["popup_inner_pad"],
        )

        settings_btn = self._popup_button(
            bottom_bar,
            text="\ue713  Settings",
            width=self.metrics["settings_button_w"],
            height=self.metrics["menu_button_h"],
            anchor="w",
            command=lambda: [
                self._launch_path("ms-settings:"),
                self._destroy_popup("start_menu"),
            ],
        )
        settings_btn.pack(
            side="left",
            padx=self.metrics["control_gap"],
            pady=self.metrics["control_gap"],
        )
        self._popup_button(
            bottom_bar,
            text="\ue7e8",
            width=self.metrics["small_icon_button"],
            height=self.metrics["small_icon_button"],
            font=self.fonts["icon_md"],
            variant="danger",
            command=lambda: subprocess.Popen(
                ["shutdown", "/s", "/t", "0"],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ),
        ).pack(
            side="right",
            padx=self.metrics["control_gap"],
            pady=self.metrics["control_gap"],
        )
        self._popup_button(
            bottom_bar,
            text="\uec46",
            width=self.metrics["small_icon_button"],
            height=self.metrics["small_icon_button"],
            font=self.fonts["icon_md"],
            variant="soft",
            command=lambda: subprocess.Popen(
                ["rundll32.exe", "powrprof.dll,SetSuspendState", "0,1,0"],
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            ),
        ).pack(
            side="right",
            padx=(0, self.metrics["control_gap"]),
            pady=self.metrics["control_gap"],
        )

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
        existing = getattr(self, "search_window", None)
        if existing is not None and existing.winfo_exists():
            existing.geometry(
                self._popup_geometry(
                    search_w,
                    search_h,
                    align="center",
                    gap=self._popup_large_gap,
                )
            )
            existing.deiconify()
            existing.wm_attributes(
                "-topmost", True, "-alpha", self._popup_target_alpha()
            )
            existing.lift()
            if hasattr(self, "search_entry") and self.search_entry.winfo_exists():
                self.search_entry.focus_force()
                self.search_entry.icursor("end")
            self.update_search_results()
            return

        self.search_window, main_frame = self._create_popup_shell(
            "search_window",
            search_w,
            search_h,
            align="center",
            gap=self._popup_large_gap,
            corner_radius=self.metrics["popup_radius"],
        )
        self._popup_header(
            main_frame,
            "\ue721",
            "Search",
            subtitle="Find and launch apps instantly",
        )
        search_area = ctk.CTkFrame(
            main_frame,
            fg_color=self.theme["surface_alt"],
            corner_radius=self.metrics["panel_radius"],
            border_width=1,
            border_color=self.theme["separator"],
        )
        search_area.pack(
            fill="x",
            padx=self.metrics["section_pad_x"],
            pady=(self.metrics["section_pad_x"], self.metrics["popup_inner_pad"]),
        )
        ctk.CTkLabel(
            search_area,
            text="\ue721",
            font=self.fonts["icon_md"],
            text_color=self.theme["text_muted"],
        ).pack(side="left", padx=(12, 0))
        self.search_entry = ctk.CTkEntry(
            search_area,
            placeholder_text="Search apps...",
            width=self.metrics["search_entry_w"],
            height=self.metrics["search_entry_h"],
            font=self.fonts["search"],
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
        self.results_scroll.pack(
            fill="both",
            expand=True,
            padx=self.metrics["popup_inner_pad"],
            pady=(0, self.metrics["popup_inner_pad"]),
        )
        self.search_entry.bind("<KeyRelease>", self._schedule_search_update)
        self.search_window.bind(
            "<Escape>", lambda e: self._destroy_popup("search_window")
        )
        self.search_entry.bind("<Return>", lambda e: self.launch_top_result())
        self._bind_popup_focus_close("search_window", anchor_widget=self.search_btn)
        self._last_search_query = None
        self._search_last_render_key = None
        self.update_search_results()
        self._fade_in(self.search_window, self._popup_target_alpha())
        self.search_window.focus_set()

    def _schedule_search_update(self, event=None):
        if event is not None and getattr(event, "keysym", "") in ("Escape", "Return"):
            return
        after_id = getattr(self, "_search_after_id", None)
        if after_id is not None:
            try:
                self.after_cancel(after_id)
            except Exception:
                pass
        self._search_after_id = self.after(45, self.update_search_results)

    def _clear_search_results(self, render_key):
        if render_key == self._search_last_render_key:
            return False
        for child in self.results_scroll.winfo_children():
            child.destroy()
        self._search_last_render_key = render_key
        return True

    def update_search_results(self, event=None):
        self._search_after_id = None
        if not self._popup_exists("search_window"):
            return
        query = self.search_entry.get()
        if query == self._last_search_query and event is not None:
            return
        self._last_search_query = query

        if len(query) < 1:
            if not self._clear_search_results(("empty",)):
                return
            ctk.CTkLabel(
                self.results_scroll,
                text="\ue721",
                font=self.fonts["icon_xl"],
                text_color=self.theme["text_faint"],
            ).pack(pady=(54, 8))
            ctk.CTkLabel(
                self.results_scroll,
                text="Start typing to search apps",
                font=self.fonts["body_bold"],
                text_color=self.theme["text_dim"],
            ).pack()
            ctk.CTkLabel(
                self.results_scroll,
                text="Press Enter to open the top result.",
                font=self.fonts["body"],
                text_color=self.theme["text_muted"],
            ).pack(pady=(4, 0))
            return
        apps = search_windows_apps(query)
        with _app_index_lock:
            index_ready = _app_index_ready
            index_building = _app_index_building
        if not apps and (not index_ready or index_building):
            _start_app_index_build()
            if self._clear_search_results(("indexing", query)):
                ctk.CTkLabel(
                    self.results_scroll,
                    text="\ue895",
                    font=self.fonts["icon_lg"],
                    text_color=self.theme["text_faint"],
                ).pack(pady=(40, 6))
                ctk.CTkLabel(
                    self.results_scroll,
                    text="Indexing apps...",
                    font=self.fonts["body"],
                    text_color=self.theme["text_muted"],
                ).pack()
            self.after(120, self.update_search_results)
            return
        if not apps:
            if not self._clear_search_results(("none", query)):
                return
            ctk.CTkLabel(
                self.results_scroll,
                text="\ue721",
                font=self.fonts["icon_lg"],
                text_color=self.theme["text_faint"],
            ).pack(pady=(40, 6))
            ctk.CTkLabel(
                self.results_scroll,
                text="No results found",
                font=self.fonts["body"],
                text_color=self.theme["text_muted"],
            ).pack()
            return
        render_key = (
            "results",
            query,
            tuple((app.get("name"), app.get("path")) for app in apps),
        )
        if not self._clear_search_results(render_key):
            return
        self._search_request += 1
        request_id = self._search_request
        for app in apps:
            try:
                app_path = app["path"]
                app_name = app["name"]
                icon = self._get_ready_ctk_icon(app_path, 28)
                result_btn = ctk.CTkButton(
                    self.results_scroll,
                    text=f"  {app_name}",
                    image=icon,
                    anchor="w",
                    height=self.metrics["menu_large_row_h"],
                    border_width=0,
                    corner_radius=self.metrics["row_radius"],
                    fg_color="transparent",
                    hover_color=self.theme["surface_hover"],
                    font=self.fonts["body_bold"],
                    text_color=self.theme["text_dim"],
                    command=lambda p=app_path: [
                        os.startfile(p),
                        self._destroy_popup("search_window"),
                    ],
                )
                result_btn.pack(fill="x", pady=2, padx=5)
                self.after_idle(
                    lambda b=result_btn, p=app_path, rid=request_id, name=app_name: (
                        self._load_button_icon_async(
                            b,
                            p,
                            28,
                            popup_attr="search_window",
                            request_attr="_search_request",
                            request_id=rid,
                            text_with_icon=f"  {name}",
                        )
                    )
                )
            except Exception as exc:
                log_error(f"update_search_results({app.get('name', 'unknown')})", exc)

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
        last_exc = None
        for attempt in range(2):
            try:
                CoInitialize()
                try:
                    enumerator = CoCreateInstance(
                        CLSID_MMDeviceEnumerator,
                        interface=IMMDeviceEnumerator,
                        clsctx=CLSCTX_ALL,
                    )
                    device = enumerator.GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
                    interface = device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, None)
                    endpoint = cast(interface, POINTER(IAudioEndpointVolume))
                    self._volume_available = True
                    return fn(endpoint)
                finally:
                    try:
                        CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                last_exc = exc
                if attempt == 0:
                    time.sleep(0.05)
        self._volume_available = False
        log_error("_with_endpoint_volume", last_exc)
        return None

    def _get_system_volume_state(self):
        def _read(endpoint):
            level = float(endpoint.GetMasterVolumeLevelScalar())
            mute = bool(endpoint.GetMute())
            return level, mute

        state = self._with_endpoint_volume(_read)
        return state if state is not None else self._volume_state_cache

    def _get_volume_and_device_name(self):
        """Read volume state and device name in a single COM session."""
        if not _AUDIO_API_AVAILABLE:
            self._volume_available = False
            return self._volume_state_cache, "Audio unavailable"
        last_exc = None
        for attempt in range(2):
            try:
                CoInitialize()
                try:
                    enumerator = CoCreateInstance(
                        CLSID_MMDeviceEnumerator,
                        interface=IMMDeviceEnumerator,
                        clsctx=CLSCTX_ALL,
                    )
                    device = enumerator.GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
                    # Read volume
                    raw_iface = device.Activate(IID_IAudioEndpointVolume, CLSCTX_ALL, None)
                    endpoint = cast(raw_iface, POINTER(IAudioEndpointVolume))
                    level = float(endpoint.GetMasterVolumeLevelScalar())
                    mute = bool(endpoint.GetMute())
                    # Read device name
                    dev_name = "Default output"
                    try:
                        store_obj = device.OpenPropertyStore(STGM_READ)
                        store = cast(store_obj, POINTER(IPropertyStore))
                        value = store.GetValue(byref(PKEY_Device_FriendlyName))
                        try:
                            if value.vt == VT_LPWSTR and value.pwszVal:
                                dev_name = str(value.pwszVal)
                        finally:
                            _PropVariantClear(byref(value))
                    except Exception:
                        pass  # device name is non-critical
                    self._volume_available = True
                    return (level, mute), dev_name
                finally:
                    try:
                        CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                last_exc = exc
                if attempt == 0:
                    time.sleep(0.05)
        self._volume_available = False
        log_error("_get_volume_and_device_name", last_exc)
        return self._volume_state_cache, "Audio unavailable"

    def _get_default_output_name(self):
        if not _AUDIO_API_AVAILABLE:
            return "Audio unavailable"
        last_exc = None
        for attempt in range(2):
            try:
                CoInitialize()
                try:
                    enumerator = CoCreateInstance(
                        CLSID_MMDeviceEnumerator,
                        interface=IMMDeviceEnumerator,
                        clsctx=CLSCTX_ALL,
                    )
                    device = enumerator.GetDefaultAudioEndpoint(E_RENDER, E_MULTIMEDIA)
                    store_obj = device.OpenPropertyStore(STGM_READ)
                    store = cast(store_obj, POINTER(IPropertyStore))
                    value = store.GetValue(byref(PKEY_Device_FriendlyName))
                    try:
                        if value.vt == VT_LPWSTR and value.pwszVal:
                            name = str(value.pwszVal)
                        else:
                            name = "Default output"
                    finally:
                        _PropVariantClear(byref(value))
                    self._volume_available = True
                    return name
                finally:
                    try:
                        CoUninitialize()
                    except Exception:
                        pass
            except Exception as exc:
                last_exc = exc
                if attempt == 0:
                    time.sleep(0.05)
        self._volume_available = False
        log_error("_get_default_output_name", last_exc)
        return "Audio unavailable"

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

    def _volume_status_text(self, level, muted):
        percent = int(round(level * 100))
        if not self._volume_available:
            return "Unavailable"
        if muted:
            return f"Muted at {percent}%"
        return f"{percent}%"

    def _sync_volume_button_icon(self, state=None):
        if state is None:
            level, muted = self._volume_state_cache
        else:
            level, muted = state
            self._volume_state_cache = (level, muted)
        try:
            self.volume_control_btn.configure(
                text=self._volume_glyph(level, muted),
                text_color=self.theme["text_faint"]
                if muted
                else self.theme["text_dim"],
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

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_volume_menu_state(self, level, muted, device_name):
        self._volume_device_cache = device_name
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
            self.volume_status_label.configure(
                text=self._volume_status_text(level, muted),
                text_color=self.theme["danger_text"]
                if muted
                else self.theme["text_dim"],
            )
        if hasattr(self, "volume_icon_label"):
            self.volume_icon_label.configure(
                text=glyph,
                text_color=self.theme["text_faint"] if muted else self.theme["accent"],
            )
        if hasattr(self, "volume_mute_btn"):
            self.volume_mute_btn.configure(
                text=("\ue74f  Unmute" if muted else "\ue74f  Mute"),
                fg_color=self.theme["surface_danger"]
                if muted
                else self.theme["surface_alt"],
                hover_color=self.theme["surface_danger_hover"]
                if muted
                else self.theme["surface_hover"],
                text_color=self.theme["danger_text"]
                if muted
                else self.theme["text_dim"],
            )
        if hasattr(self, "volume_quick_buttons"):
            active_target = min((0, 25, 50, 75, 100), key=lambda p: abs(p - percent))
            for pct, btn in self.volume_quick_buttons.items():
                is_active = (
                    pct == active_target and not muted and self._volume_available
                )
                btn.configure(
                    fg_color=self.theme["surface_active"]
                    if is_active
                    else "transparent",
                    text_color=self.theme["accent"]
                    if is_active
                    else self.theme["text_dim"],
                    border_width=1 if is_active else 0,
                    border_color=self.theme["surface_active_border"],
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

        def _worker():
            (level, muted), device_name = self._get_volume_and_device_name()
            self._post_ui(
                lambda: self._apply_volume_menu_state(level, muted, device_name)
            )

        threading.Thread(target=_worker, daemon=True).start()

    def _refresh_volume_menu(self):
        if not hasattr(self, "volume_menu") or not self.volume_menu.winfo_exists():
            return

        (level, muted), device_name = self._get_volume_and_device_name()
        self._apply_volume_menu_state(level, muted, device_name)

    def _apply_volume_scalar_async(self, scalar):
        level, muted = self._volume_state_cache
        muted = False if scalar > 0 else muted
        self._volume_state_cache = (scalar, muted)
        self._apply_volume_menu_state(scalar, muted, self._volume_device_cache)
        self._pending_volume_scalar = scalar
        if self._volume_set_inflight:
            return
        self._volume_set_inflight = True

        def _worker():
            while True:
                value = self._pending_volume_scalar
                self._pending_volume_scalar = None
                if value is None:
                    break
                self._set_system_volume(value)
                if value > 0:
                    self._set_system_mute(False)
                if self._pending_volume_scalar is None:
                    break

            def _apply():
                self._volume_set_inflight = False
                if self._pending_volume_scalar is not None:
                    self._apply_volume_scalar_async(self._pending_volume_scalar)
                else:
                    self._refresh_volume_menu_async()

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_mute_async(self, muted):
        level, _old_muted = self._volume_state_cache
        self._volume_state_cache = (level, bool(muted))
        self._apply_volume_menu_state(level, bool(muted), self._volume_device_cache)
        self._pending_mute_value = bool(muted)
        if self._mute_set_inflight:
            return
        self._mute_set_inflight = True

        def _worker():
            while True:
                value = self._pending_mute_value
                self._pending_mute_value = None
                if value is None:
                    break
                self._set_system_mute(value)
                if self._pending_mute_value is None:
                    break

            def _apply():
                self._mute_set_inflight = False
                if self._pending_mute_value is not None:
                    self._apply_mute_async(self._pending_mute_value)
                else:
                    self._refresh_volume_menu_async()

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_volume_slider(self, value):
        if self._volume_slider_busy:
            return
        scalar = max(0.0, min(1.0, float(value) / 100.0))
        self._apply_volume_scalar_async(scalar)

    def _toggle_system_mute(self):
        level, muted = self._volume_state_cache
        self._apply_mute_async(not muted)

    def _set_quick_volume(self, percent):
        self._on_volume_slider(percent)

    def _nudge_volume(self, delta_percent):
        level, _muted = self._volume_state_cache
        next_percent = int(round(level * 100)) + int(delta_percent)
        self._on_volume_slider(max(0, min(100, next_percent)))

    def _on_volume_wheel(self, event):
        if not self._popup_exists("volume_menu"):
            return
        step = 5 if getattr(event, "delta", 0) > 0 else -5
        self._nudge_volume(step)
        return "break"

    def _open_file_location(self, path):
        try:
            subprocess.Popen(["explorer", f"/select,{path}"])
        except Exception as exc:
            log_error(f"_open_file_location({path})", exc)

    def _get_cached_background_icon(self, path):
        norm = os.path.normcase(path)
        if norm not in self._background_icon_cache:
            self._background_icon_cache[norm] = get_icon_from_exe(path)
        img = self._background_icon_cache.get(norm)
        return ctk.CTkImage(img, img, size=(20, 20)) if img else None

    def _clone_background_apps(self, bg_apps):
        return [{**app, "pids": list(app.get("pids", []))} for app in (bg_apps or [])]

    def _background_apps_signature(self, bg_apps):
        return tuple(
            (
                os.path.normcase(app.get("path", "")),
                int(app.get("count", 0)),
            )
            for app in (bg_apps or [])
        )

    def _set_background_apps_cache(self, bg_apps):
        self._background_apps_cache = self._clone_background_apps(bg_apps)
        self._background_apps_cache_signature = self._background_apps_signature(bg_apps)
        self._background_apps_cache_time = time.monotonic()

    def _prime_background_icons(self, bg_apps):
        for app in (bg_apps or [])[:15]:
            path = app.get("path")
            if not path:
                continue
            norm = os.path.normcase(path)
            if (
                norm in self._background_icon_cache
                or norm in self._background_icon_load_inflight
            ):
                continue
            self._background_icon_load_inflight.add(norm)

            def _worker(icon_path=path, icon_key=norm):
                img = get_icon_from_exe(icon_path)

                def _apply():
                    self._background_icon_load_inflight.discard(icon_key)
                    self._background_icon_cache[icon_key] = img or False

                self._post_ui(_apply)

            threading.Thread(target=_worker, daemon=True).start()

    def _render_background_apps_loading(self, message="Loading background apps..."):
        grid = getattr(self, "background_apps_grid", None)
        if grid is None or not grid.winfo_exists():
            return
        for child in grid.winfo_children():
            child.destroy()
        ctk.CTkLabel(
            grid,
            text="\ue895",
            font=self.fonts["icon_lg"],
            text_color=self.theme["text_faint"],
        ).pack(pady=(34, 6))
        ctk.CTkLabel(
            grid,
            text=message,
            font=self.fonts["label"],
            text_color=self.theme["text_muted"],
        ).pack()

    def _apply_background_tile_icon(self, tile, path, request_id, img):
        norm = os.path.normcase(path)
        self._background_icon_cache[norm] = img or False
        if (
            request_id != self._background_apps_request
            or not self._popup_exists("background_apps_menu")
            or not tile.winfo_exists()
        ):
            return
        if img:
            icon = ctk.CTkImage(img, img, size=(20, 20))
            tile.configure(image=icon, text="")
            tile._icon_ref = icon

    def _load_background_tile_icon_async(self, tile, path, request_id):
        norm = os.path.normcase(path)
        if norm in self._background_icon_cache:
            cached = self._background_icon_cache.get(norm)
            if cached:
                self._apply_background_tile_icon(tile, path, request_id, cached)
            return
        if norm in self._background_icon_load_inflight:
            self.after(
                35,
                lambda: self._load_background_tile_icon_async(tile, path, request_id),
            )
            return

        self._background_icon_load_inflight.add(norm)

        def _worker():
            img = get_icon_from_exe(path)

            def _apply():
                self._background_icon_load_inflight.discard(norm)
                self._apply_background_tile_icon(tile, path, request_id, img)

            self._post_ui(_apply)

        threading.Thread(target=_worker, daemon=True).start()

    def _render_background_apps_contents(self, bg_apps, request_id=None):
        grid = getattr(self, "background_apps_grid", None)
        if grid is None or not grid.winfo_exists():
            return
        if request_id is None:
            request_id = self._background_apps_request

        for child in grid.winfo_children():
            child.destroy()

        if not bg_apps:
            ctk.CTkLabel(
                grid,
                text="No background apps found",
                font=self.fonts["label"],
                text_color=self.theme["text_muted"],
            ).pack(anchor="w", padx=6, pady=(10, 2))
            return

        for col in range(5):
            grid.grid_columnconfigure(col, weight=1)
        for idx, app in enumerate(bg_apps):
            tile = ctk.CTkButton(
                grid,
                text=(app["name"][:1] or "?").upper(),
                width=self.metrics["tile_size"],
                height=self.metrics["tile_size"],
                fg_color=self.theme["surface_alt"],
                hover_color=self.theme["surface_hover"],
                corner_radius=self.metrics["tile_radius"],
                font=self.fonts["label_bold"],
                text_color=self.theme["text_muted"],
                border_width=1,
                border_color=self.theme["separator"],
                anchor="center",
                border_spacing=0,
                command=lambda p=app["path"]: self._launch_background_app(p),
            )
            tile.grid(
                row=idx // 5,
                column=idx % 5,
                padx=self.metrics["grid_gap"],
                pady=self.metrics["grid_gap"],
                sticky="nsew",
            )
            tooltip = app["name"]
            if app["count"] > 1:
                tooltip += f" ({app['count']})"
            tile.bind(
                "<Enter>", lambda e, w=tile, t=tooltip: self._schedule_tooltip(w, t)
            )
            tile.bind("<Leave>", lambda e, w=tile: self._schedule_tooltip_hide(w))
            tile.bind("<ButtonPress-1>", lambda e: self._hide_tooltip(), add="+")
            tile.bind(
                "<Button-3>",
                lambda e, w=tile, a=app: self._show_background_app_menu(w, a),
                add="+",
            )
            self._load_background_tile_icon_async(tile, app["path"], request_id)

    def _apply_background_apps_update(self, request_id, bg_apps):
        self._background_apps_refresh_inflight = False
        if request_id != self._background_apps_request:
            return

        next_signature = self._background_apps_signature(bg_apps)
        cache_changed = self._background_apps_cache_signature != next_signature
        self._set_background_apps_cache(bg_apps)
        self._prime_background_icons(bg_apps)

        if not self._popup_exists("background_apps_menu"):
            return

        if cache_changed:
            self._render_background_apps_contents(bg_apps, request_id)

    def _load_background_apps_async(self, request_id=None):
        if self._background_apps_refresh_inflight:
            return
        if request_id is None:
            self._background_apps_request += 1
            request_id = self._background_apps_request
        self._background_apps_refresh_inflight = True

        def _worker():
            try:
                apps = get_background_apps(limit=15)
            except Exception as exc:
                log_error("_load_background_apps_async", exc)
                apps = []
            self._post_ui(lambda: self._apply_background_apps_update(request_id, apps))

        threading.Thread(target=_worker, daemon=True).start()

    def _launch_background_app(self, path):
        self._hide_tooltip()
        self._launch_path(path)
        self._destroy_popup("background_apps_menu")

    def _refresh_background_apps_menu(self):
        try:
            if self._popup_exists("background_apps_menu"):
                self._destroy_popup("background_apps_menu")
                self.toggle_background_apps()
        except Exception as exc:
            log_error("_refresh_background_apps_menu", exc)

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
            corner_radius=self.metrics["panel_radius"],
        )

        self._popup_button(
            frame,
            text="Open app",
            height=self.metrics["menu_button_h"],
            anchor="w",
            command=lambda p=app["path"]: [
                self._launch_background_app(p),
                menu.destroy(),
            ],
        ).pack(fill="x", padx=6, pady=(6, 2))
        self._popup_button(
            frame,
            text="Open file location",
            height=self.metrics["menu_button_h"],
            anchor="w",
            command=lambda p=app["path"]: [self._open_file_location(p), menu.destroy()],
        ).pack(fill="x", padx=6, pady=2)
        self._popup_button(
            frame,
            text="End task",
            height=self.metrics["menu_button_h"],
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

        menu_w, menu_h = 310, 292
        self.volume_menu, f = self._create_popup_shell(
            "volume_menu",
            menu_w,
            menu_h,
            anchor_widget=self.volume_control_btn,
            align="right",
            corner_radius=self.metrics["popup_radius"],
        )

        self.volume_icon_label, self.volume_status_label = self._popup_header(
            f,
            "\ue995",
            "Volume",
            subtitle="Output level",
            value="0%",
        )

        output_card = ctk.CTkFrame(
            f,
            fg_color=self.theme["surface_alt"],
            corner_radius=self.metrics["row_radius"],
            border_width=1,
            border_color=self.theme["separator"],
        )
        output_card.pack(
            fill="x",
            padx=self.metrics["control_pad_x"],
            pady=(0, self.metrics["popup_inner_pad"]),
        )
        output_card.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            output_card,
            text="\ue7f5",
            font=self.fonts["icon_sm"],
            text_color=self.theme["accent"],
        ).grid(row=0, column=0, padx=(10, 8), pady=9)
        self.volume_device_label = ctk.CTkLabel(
            output_card,
            text="Default output",
            font=self.fonts["caption_bold"],
            text_color=self.theme["text_dim"],
            anchor="w",
        )
        self.volume_device_label.grid(row=0, column=1, sticky="ew", padx=(0, 10))

        slider_row = ctk.CTkFrame(f, fg_color="transparent")
        slider_row.pack(fill="x", padx=self.metrics["control_pad_x"], pady=(0, 12))
        down_btn = self._popup_button(
            slider_row,
            text="-",
            width=32,
            height=32,
            variant="soft",
            font=self.fonts["label_bold"],
            command=lambda: self._nudge_volume(-5),
        )
        down_btn.pack(side="left", padx=(0, 8))
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
        up_btn = self._popup_button(
            slider_row,
            text="+",
            width=32,
            height=32,
            variant="soft",
            font=self.fonts["label_bold"],
            command=lambda: self._nudge_volume(5),
        )
        up_btn.pack(side="left", padx=(8, 0))

        action_row = ctk.CTkFrame(f, fg_color="transparent")
        action_row.pack(fill="x", padx=self.metrics["control_pad_x"], pady=(0, 8))
        self.volume_mute_btn = self._popup_button(
            action_row,
            text="\ue74f  Mute",
            height=self.metrics["menu_button_h"],
            variant="soft",
            command=self._toggle_system_mute,
        )
        self.volume_mute_btn.pack(fill="x")

        quick_row = ctk.CTkFrame(f, fg_color="transparent")
        quick_row.pack(
            fill="x",
            padx=self.metrics["control_pad_x"],
            pady=(0, self.metrics["popup_inner_pad"]),
        )
        for idx in range(5):
            quick_row.grid_columnconfigure(idx, weight=1, uniform="volume_quick")
        self.volume_quick_buttons = {}
        for idx, pct in enumerate((0, 25, 50, 75, 100)):
            btn = self._popup_button(
                quick_row,
                text=f"{pct}%",
                width=1,
                height=self.metrics["menu_button_h"],
                variant="ghost",
                command=lambda p=pct: self._set_quick_volume(p),
            )
            btn.grid(
                row=0,
                column=idx,
                sticky="ew",
                padx=(0, self.metrics["grid_gap"]) if idx < 4 else (0, 0),
            )
            self.volume_quick_buttons[pct] = btn

        self.volume_hint_label = ctk.CTkLabel(
            f,
            text="Master output",
            font=self.fonts["caption"],
            text_color=self.theme["text_faint"],
        )
        self.volume_hint_label.pack(
            anchor="w", padx=self.metrics["control_pad_x"], pady=(2, 0)
        )

        cached_level, cached_muted = self._volume_state_cache
        self._apply_volume_menu_state(
            cached_level, cached_muted, self._volume_device_cache
        )
        self._bind_popup_focus_close(
            "volume_menu", anchor_widget=self.volume_control_btn
        )
        self._fade_in(self.volume_menu, self._popup_target_alpha())
        self.volume_menu.focus_set()
        self.volume_menu.bind("<MouseWheel>", self._on_volume_wheel, add="+")
        f.bind("<MouseWheel>", self._on_volume_wheel, add="+")
        self.after_idle(self._refresh_volume_menu_async)

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
            corner_radius=self.metrics["popup_radius"],
        )
        self._popup_header(
            panel,
            "\uec8f",
            "Background apps",
            subtitle="Quiet processes currently running",
        )
        self._popup_divider(panel, pady=(0, 8))

        self.background_apps_grid = ctk.CTkScrollableFrame(
            panel, fg_color="transparent", corner_radius=0
        )
        self.background_apps_grid.pack(
            fill="both",
            expand=True,
            padx=self.metrics["popup_inner_pad"],
            pady=(0, self.metrics["popup_inner_pad"]),
        )
        self._background_apps_request += 1
        request_id = self._background_apps_request
        if self._background_apps_cache is None:
            self._render_background_apps_loading("Loading background apps...")
        else:
            self._render_background_apps_contents(
                self._background_apps_cache, request_id
            )
        self._bind_popup_focus_close(
            "background_apps_menu",
            anchor_widget=self.background_apps_btn,
            allow=("background_app_context_menu",),
        )
        self._fade_in(self.background_apps_menu, self._popup_target_alpha())
        self.background_apps_menu.focus_set()
        cache_age = time.monotonic() - self._background_apps_cache_time
        if (
            self._background_apps_cache is None
            or cache_age > self._background_apps_cache_ttl
        ):
            self.after(
                250, lambda rid=request_id: self._load_background_apps_async(rid)
            )

    def toggle_control_center(self):
        if self.edit_mode:
            return
        if self._popup_exists("tray_menu"):
            self._destroy_popup("tray_menu")
            return
        self._close_all_popups(exclude=("tray_menu",))

        menu_w, menu_h = 300, 470
        self.tray_menu, f = self._create_popup_shell(
            "tray_menu",
            menu_w,
            menu_h,
            anchor_widget=self.tray_btn,
            align="right",
            corner_radius=self.metrics["popup_radius"],
        )
        self._popup_header(
            f,
            "\ue782",
            "WinBar",
            subtitle="Display, position, and layout",
        )
        self._popup_divider(f, pady=(0, 8))

        # Opacity
        ctk.CTkLabel(
            f,
            text="Opacity",
            font=self.fonts["caption"],
            text_color=self.theme["text_muted"],
        ).pack(anchor="w", padx=self.metrics["control_pad_x"])
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
        opacity_slider.pack(
            pady=(2, self.metrics["popup_inner_pad"]),
            padx=self.metrics["control_pad_x"],
            fill="x",
        )

        # Position
        ctk.CTkLabel(
            f,
            text="Position",
            font=self.fonts["caption"],
            text_color=self.theme["text_muted"],
        ).pack(anchor="w", padx=self.metrics["control_pad_x"])
        pos_menu = ctk.CTkOptionMenu(
            f,
            values=["Top", "Bottom"],
            command=self.change_position,
            width=self.metrics["option_menu_w"],
            fg_color=self.theme["surface_alt"],
            button_color=self.theme["surface_hover"],
            button_hover_color=self.theme["accent_hover"],
            dropdown_fg_color=self.theme["surface_alt"],
            dropdown_hover_color=self.theme["surface_hover"],
            text_color=self.theme["text_dim"],
            font=self.fonts["caption_bold"],
            dropdown_font=self.fonts["caption"],
        )
        pos_menu.set(self.config.get("position", "Top"))
        pos_menu.pack(
            pady=(2, self.metrics["popup_inner_pad"]),
            padx=self.metrics["control_pad_x"],
        )

        # Hover tooltips
        ctk.CTkLabel(
            f,
            text="Hover Tooltips",
            font=self.fonts["caption"],
            text_color=self.theme["text_muted"],
        ).pack(anchor="w", padx=self.metrics["control_pad_x"])
        tooltip_menu = ctk.CTkOptionMenu(
            f,
            values=["On", "Off"],
            command=self.change_hover_tooltips,
            width=self.metrics["option_menu_w"],
            fg_color=self.theme["surface_alt"],
            button_color=self.theme["surface_hover"],
            button_hover_color=self.theme["accent_hover"],
            dropdown_fg_color=self.theme["surface_alt"],
            dropdown_hover_color=self.theme["surface_hover"],
            text_color=self.theme["text_dim"],
            font=self.fonts["caption_bold"],
            dropdown_font=self.fonts["caption"],
        )
        tooltip_menu.set("On" if self.config.get("hover_tooltips", True) else "Off")
        tooltip_menu.pack(
            pady=(2, self.metrics["popup_inner_pad"]),
            padx=self.metrics["control_pad_x"],
        )

        # Actions
        self._popup_button(
            f,
            text="Edit Layout",
            height=self.metrics["menu_button_h"],
            variant="primary",
            command=self.toggle_edit_mode,
        ).pack(
            padx=self.metrics["control_pad_x"],
            pady=(0, self.metrics["control_gap"]),
            fill="x",
        )
        self._popup_button(
            f,
            text="Quit WinBar",
            height=self.metrics["menu_button_h"],
            variant="danger",
            command=self.safe_exit,
        ).pack(
            padx=self.metrics["control_pad_x"],
            pady=(0, self.metrics["divider_pad_x"]),
            fill="x",
        )
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

    def change_hover_tooltips(self, value):
        enabled = value == "On"
        self.config["hover_tooltips"] = enabled
        if not enabled:
            self._hide_tooltip()
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
            y_pos = sh - h - self._edge_gap
            edge = ABE_BOTTOM
            appbar_pad = self._edge_gap
        else:
            y_pos = 0
            edge = ABE_TOP
            appbar_pad = 0

        self.geo_string = f"{w}x{h}+{x_pos}+{y_pos}"
        self.geometry(self.geo_string)
        self.update()

        HWND_TOPMOST = -1
        my_hwnd = ctypes.windll.user32.GetAncestor(self.winfo_id(), GA_ROOT)
        ctypes.windll.user32.SetWindowPos(
            my_hwnd, HWND_TOPMOST, x_pos, y_pos, w, h, 0x0040
        )
        unregister_appbar()
        register_appbar(self.winfo_id(), h, appbar_pad, edge)
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
            corner_radius=self.metrics["popup_radius_compact"],
        )
        self._popup_header(
            main_frame,
            "\ue712",
            "More apps",
            subtitle="Running apps beyond the bar",
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
                height=self.metrics["menu_row_h"],
                fg_color="transparent",
                hover_color=self.theme["surface_hover"],
                corner_radius=self.metrics["row_radius"],
                font=self.fonts["body_bold"],
                text_color=self.theme["text_dim"],
                command=lambda g=group_key: (
                    [self.focus_group(g), self._destroy_popup("overflow_menu")]
                    if g
                    else None
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
                fg_color=self.theme["bar_bg"],
                border_color=self.theme["warning"],
                border_width=2,
            )
            self.active_window_label.configure(
                text="\ue70f  Edit Mode  -  drag to rearrange",
                text_color=self.theme["warning"],
            )
            self._set_cursor_recursive(self.pill_frame, "fleur")
            # Show draggable zone borders on all widgets
            for w in self.widget_map.values():
                try:
                    w.configure(border_width=1, border_color=self.theme["warning_dark"])
                except Exception:
                    pass
            self.edit_done_btn.pack(in_=self.right_wing, side="right", padx=(4, 2))
        else:
            self._cancel_drag()  # clean up any in-progress drag
            self.pill_frame.configure(
                fg_color=self.theme["bar_bg"],
                border_color=self.theme["bar_border"],
                border_width=1,
            )
            self.active_window_label.configure(
                text="", text_color=self.theme["text_muted"]
            )
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
                container = self.active_app_buttons[active_group]
                if container.winfo_exists():
                    self._set_group_button_state(container, active=True)
            self.save_layout()

    def save_layout(self):
        with open(self.config_path, "w") as f:
            json.dump(self.config, f, indent=4)

    # Drag-and-drop helpers

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
        ghost.configure(bg=self.theme["warning"])
        outer = tk.Frame(
            ghost,
            bg=self.theme["warning"],
            padx=self.metrics["popup_inner_pad"],
            pady=5,
        )
        outer.pack()
        tk.Label(
            outer,
            text=label,
            font=self.fonts["caption_bold"],
            fg=self.theme["warning_text"],
            bg=self.theme["warning"],
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
                w.configure(border_width=1, border_color=self.theme["warning_dark"])
            except Exception:
                pass

    # Drag event handlers

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
                src.configure(border_width=2, border_color=self.theme["warning"])
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
                    pw.configure(
                        border_width=1, border_color=self.theme["warning_dark"]
                    )
                except Exception:
                    pass

        # Highlight new drop target in light blue
        if target and target != source:
            tw = self.widget_map.get(target)
            if tw:
                try:
                    tw.configure(border_width=2, border_color=self.theme["drop_target"])
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
        if getattr(self, "_exiting", False):
            return
        self._exiting = True

        after_id = getattr(self, "_ui_queue_after_id", None)
        if after_id is not None:
            try:
                self.after_cancel(after_id)
            except Exception:
                pass
            self._ui_queue_after_id = None
        after_id = getattr(self, "_taskbar_guard_after_id", None)
        if after_id is not None:
            try:
                self.after_cancel(after_id)
            except Exception:
                pass
            self._taskbar_guard_after_id = None

        stop_system_tray()
        set_taskbar_visibility(True)
        unregister_appbar()
        self.quit()
        self.destroy()


if __name__ == "__main__":
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()

    set_taskbar_visibility(False)
    reset_work_area()
    atexit.register(unregister_appbar)
    atexit.register(lambda: set_taskbar_visibility(True))
    atexit.register(stop_system_tray)

    app = FloatingBar()
    app.protocol("WM_DELETE_WINDOW", app.safe_exit)
    app.update_idletasks()

    root = ctypes.windll.user32.GetAncestor(app.winfo_id(), GA_ROOT)
    style = ctypes.windll.user32.GetWindowLongW(root, GWL_EXSTYLE)
    ctypes.windll.user32.SetWindowLongW(
        root, GWL_EXSTYLE, (style & ~WS_EX_APPWINDOW) | WS_EX_TOOLWINDOW
    )

    edge = ABE_BOTTOM if app.config.get("position", "Top") == "Bottom" else ABE_TOP
    register_appbar(
        app.winfo_id(),
        int(app.config.get("height", 50)),
        app._edge_gap if edge == ABE_BOTTOM else 0,
        edge,
    )

    app.geometry(app.geo_string)
    app.update_time()
    app._poll_foreground()
    app.update_open_apps()
    threading.Thread(target=install_to_startup, daemon=True).start()
    threading.Thread(target=setup_system_tray, args=(app,), daemon=True).start()
    app.mainloop()
