"""
마우스 움직임 제어 (실험) — Alt+1 또는 설정▸실험 버튼으로 토글.

ON: 사용자의 실제 마우스 이동을 Windows 저수준 훅(WH_MOUSE_LL)으로 차단하되,
    1분 주기로 1초 동안만 이동을 허용(그 순간 커서를 1px 이동).
상태는 메모리에만 보관(저장 없음). 서버 프로세스가 종료되면 Windows 가 훅을
자동 해제하므로, 창을 닫거나 Alt+1 로 언제든 해제할 수 있다.
"""
import os
import threading
import time

_CYCLE_SEC = 60.0   # 1분 주기
_ALLOW_SEC = 1.0    # 1초만 이동 허용

_lock = threading.Lock()
_state = {"active": False, "supported": os.name == "nt"}
_pending_toggle = False
_allow_now = False
_nudge_dir = 1

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    _u = ctypes.windll.user32
    _k = ctypes.windll.kernel32

    WH_MOUSE_LL = 14
    WM_MOUSEMOVE = 0x0200
    LLMHF_INJECTED = 0x00000001
    VK_MENU = 0x12
    VK_1 = 0x31
    VK_NUMPAD1 = 0x61

    ULONG_PTR = wintypes.WPARAM
    LRESULT = ctypes.c_ssize_t
    HOOKPROC = ctypes.CFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

    class _MSLL(ctypes.Structure):
        _fields_ = [("pt", wintypes.POINT), ("mouseData", wintypes.DWORD),
                    ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("dwExtraInfo", ULONG_PTR)]

    _u.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD]
    _u.SetWindowsHookExW.restype = ctypes.c_void_p
    _u.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
    _u.CallNextHookEx.restype = LRESULT
    _u.GetMessageW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT]
    _u.GetAsyncKeyState.argtypes = [ctypes.c_int]
    _u.GetAsyncKeyState.restype = ctypes.c_short
    _u.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
    _u.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
    _k.GetModuleHandleW.restype = wintypes.HMODULE
    _k.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]

    def _get_cursor():
        pt = wintypes.POINT()
        _u.GetCursorPos(ctypes.byref(pt))
        return pt.x, pt.y

    def _set_cursor(x, y):
        _u.SetCursorPos(int(x), int(y))

    def _key_down(vk):
        return (_u.GetAsyncKeyState(vk) & 0x8000) != 0

    def _ll_proc(nCode, wParam, lParam):
        try:
            if nCode == 0 and _state["active"] and not _allow_now and wParam == WM_MOUSEMOVE:
                ms = ctypes.cast(lParam, ctypes.POINTER(_MSLL)).contents
                if not (ms.flags & LLMHF_INJECTED):
                    return 1   # 사용자의 실제 이동 차단
        except Exception:
            pass
        return _u.CallNextHookEx(None, nCode, wParam, lParam)

    _hook_cb = HOOKPROC(_ll_proc)
    _hook_handle = None

    def _hook_thread():
        global _hook_handle
        _hook_handle = _u.SetWindowsHookExW(WH_MOUSE_LL, _hook_cb, _k.GetModuleHandleW(None), 0)
        if not _hook_handle:
            return
        msg = wintypes.MSG()
        while _u.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            _u.TranslateMessage(ctypes.byref(msg))
            _u.DispatchMessageW(ctypes.byref(msg))

    threading.Thread(target=_hook_thread, daemon=True, name="exp-hook").start()
else:
    def _get_cursor():
        return (0, 0)

    def _set_cursor(x, y):
        pass

    def _key_down(vk):
        return False


def _nudge():
    """이동 허용 창에서 커서를 1px 이동(좌우 번갈아)."""
    global _nudge_dir
    x, y = _get_cursor()
    _set_cursor(x + _nudge_dir, y)
    _nudge_dir = -_nudge_dir


def _controller():
    global _pending_toggle, _allow_now
    combo_was = False
    cycle_start = None
    while True:
        time.sleep(0.03)
        if os.name != "nt":
            continue
        combo = _key_down(VK_MENU) and (_key_down(VK_1) or _key_down(VK_NUMPAD1))
        edge = combo and not combo_was
        combo_was = combo

        with _lock:
            pending = _pending_toggle
            _pending_toggle = False
            active = _state["active"]

        if edge or pending:
            active = not active
            with _lock:
                _state["active"] = active
            _allow_now = False
            cycle_start = time.time() if active else None
            continue

        if not active:
            continue
        now = time.time()
        if cycle_start is None:
            cycle_start = now
        elapsed = (now - cycle_start) % _CYCLE_SEC
        want_allow = elapsed < _ALLOW_SEC
        if want_allow and not _allow_now:
            _allow_now = True
            _nudge()                 # 1초 창 열릴 때 1px 이동
        elif not want_allow and _allow_now:
            _allow_now = False


threading.Thread(target=_controller, daemon=True, name="exp-ctrl").start()


# ─── 공개 API ────────────────────────────────────────────────────────
def request_toggle():
    global _pending_toggle
    with _lock:
        _pending_toggle = True
    return {"requested": True}


def force_off():
    global _pending_toggle
    with _lock:
        if _state["active"]:
            _pending_toggle = True
    return {"requested": True}


def get_status():
    with _lock:
        return {"active": _state["active"], "supported": _state["supported"]}
