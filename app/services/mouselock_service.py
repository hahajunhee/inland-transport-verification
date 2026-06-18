"""
마우스 고정(잠금) 서비스 — Alt+1 글로벌 핫키로 토글.

ON 상태:
  - 사용자의 실제(하드웨어) 마우스 이동을 Windows 저수준 훅(WH_MOUSE_LL)으로 차단
    → 유저가 아무리 흔들어도 커서가 움직이지 않음. (주입된 이동은 통과시켜 우리 nudge 는 동작)
  - 1~3분 랜덤 간격으로 프로그램이 커서를 딱 1px 이동(활동 흔적/완전 정지 방지)
OFF 전환: Alt+1 재입력, 또는 30분 경과 시 자동 해제.

안전: 상태는 메모리에만 보관(재시작 시 off). 서버 프로세스가 종료되면 Windows 가
훅을 자동 해제하므로, 문제가 생기면 서버 창을 닫으면 마우스가 즉시 정상화된다.
"""
import os
import threading
import time
import random

_RUN_LIMIT_SEC = 30 * 60   # 30분 자동 해제
_NUDGE_MIN = 60            # 최소 1분
_NUDGE_MAX = 180           # 최대 3분
_TICK = 0.03              # 컨트롤러 폴링 주기(초)

_lock = threading.Lock()
_state = {
    "active": False,
    "activated_ts": None,
    "auto_off_ts": None,
    "nudge_count": 0,
    "last_nudge_ts": None,
    "supported": os.name == "nt",
}
_pending_toggle = False     # API 수동 토글 요청 플래그
_nudging = False            # nudge 중에는 훅이 이동을 통과시킴
_nudge_dir = 1
_next_nudge_ts = None

# ─── Windows 전용 ctypes 설정 ─────────────────────────────────────────
if os.name == "nt":
    import ctypes
    from ctypes import wintypes

    _user32 = ctypes.windll.user32
    _kernel32 = ctypes.windll.kernel32

    WH_MOUSE_LL = 14
    WM_MOUSEMOVE = 0x0200
    LLMHF_INJECTED = 0x00000001
    VK_MENU = 0x12   # Alt
    VK_1 = 0x31
    VK_NUMPAD1 = 0x61

    ULONG_PTR = wintypes.WPARAM
    LRESULT = ctypes.c_ssize_t
    HOOKPROC = ctypes.CFUNCTYPE(LRESULT, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)

    class MSLLHOOKSTRUCT(ctypes.Structure):
        _fields_ = [("pt", wintypes.POINT), ("mouseData", wintypes.DWORD),
                    ("flags", wintypes.DWORD), ("time", wintypes.DWORD),
                    ("dwExtraInfo", ULONG_PTR)]

    # 64비트 안전을 위해 핸들/포인터형 명시
    _user32.SetWindowsHookExW.argtypes = [ctypes.c_int, HOOKPROC, wintypes.HINSTANCE, wintypes.DWORD]
    _user32.SetWindowsHookExW.restype = ctypes.c_void_p
    _user32.CallNextHookEx.argtypes = [ctypes.c_void_p, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM]
    _user32.CallNextHookEx.restype = LRESULT
    _user32.UnhookWindowsHookEx.argtypes = [ctypes.c_void_p]
    _user32.GetMessageW.argtypes = [ctypes.c_void_p, wintypes.HWND, wintypes.UINT, wintypes.UINT]
    _user32.GetAsyncKeyState.argtypes = [ctypes.c_int]
    _user32.GetAsyncKeyState.restype = ctypes.c_short
    _user32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
    _user32.GetCursorPos.argtypes = [ctypes.POINTER(wintypes.POINT)]
    _kernel32.GetModuleHandleW.restype = wintypes.HMODULE
    _kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]

    def _get_cursor():
        pt = wintypes.POINT()
        _user32.GetCursorPos(ctypes.byref(pt))
        return pt.x, pt.y

    def _set_cursor(x, y):
        _user32.SetCursorPos(int(x), int(y))

    def _key_down(vk):
        return (_user32.GetAsyncKeyState(vk) & 0x8000) != 0

    def _ll_mouse_proc(nCode, wParam, lParam):
        try:
            if nCode == 0 and _state["active"] and wParam == WM_MOUSEMOVE:
                ms = ctypes.cast(lParam, ctypes.POINTER(MSLLHOOKSTRUCT)).contents
                injected = bool(ms.flags & LLMHF_INJECTED)
                # 우리 nudge(주입/_nudging) 는 통과, 사용자의 실제 이동만 차단
                if not injected and not _nudging:
                    return 1
        except Exception:
            pass
        return _user32.CallNextHookEx(None, nCode, wParam, lParam)

    _hook_proc = HOOKPROC(_ll_mouse_proc)   # GC 방지용 전역 보관
    _hook_handle = None

    def _hook_thread():
        global _hook_handle
        hmod = _kernel32.GetModuleHandleW(None)
        _hook_handle = _user32.SetWindowsHookExW(WH_MOUSE_LL, _hook_proc, hmod, 0)
        if not _hook_handle:
            return
        msg = wintypes.MSG()
        while _user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
            _user32.TranslateMessage(ctypes.byref(msg))
            _user32.DispatchMessageW(ctypes.byref(msg))

    threading.Thread(target=_hook_thread, daemon=True, name="mouselock-hook").start()
else:
    def _get_cursor():
        return (0, 0)

    def _set_cursor(x, y):
        pass

    def _key_down(vk):
        return False


# ─── 상태 전환 (컨트롤러 스레드에서만 호출) ──────────────────────────────
def _activate():
    global _next_nudge_ts, _nudge_dir
    now = time.time()
    _nudge_dir = 1
    _next_nudge_ts = now + random.uniform(_NUDGE_MIN, _NUDGE_MAX)
    with _lock:
        _state["active"] = True
        _state["activated_ts"] = now
        _state["auto_off_ts"] = now + _RUN_LIMIT_SEC


def _deactivate():
    with _lock:
        _state["active"] = False
        _state["activated_ts"] = None
        _state["auto_off_ts"] = None


def _do_nudge():
    """커서를 1px 이동(주입). 훅이 통과시키도록 _nudging 창을 잠깐 연다."""
    global _nudging, _next_nudge_ts, _nudge_dir
    x, y = _get_cursor()
    _nudging = True
    try:
        _set_cursor(x + _nudge_dir, y)
        time.sleep(0.012)   # 이동 이벤트가 훅을 통과할 시간
    finally:
        _nudging = False
    _nudge_dir = -_nudge_dir
    now = time.time()
    _next_nudge_ts = now + random.uniform(_NUDGE_MIN, _NUDGE_MAX)
    with _lock:
        _state["nudge_count"] += 1
        _state["last_nudge_ts"] = now


def _controller():
    global _pending_toggle
    combo_was_down = False
    while True:
        time.sleep(_TICK)
        if os.name != "nt":
            continue
        # Alt+1 엣지 감지
        combo = _key_down(VK_MENU) and (_key_down(VK_1) or _key_down(VK_NUMPAD1))
        hotkey_edge = combo and not combo_was_down
        combo_was_down = combo

        with _lock:
            pending = _pending_toggle
            _pending_toggle = False
            active = _state["active"]
            auto_off = _state["auto_off_ts"]

        if hotkey_edge or pending:
            if active:
                _deactivate()
            else:
                _activate()
            continue

        if not active:
            continue
        now = time.time()
        if auto_off and now >= auto_off:   # 30분 자동 해제
            _deactivate()
            continue
        if _next_nudge_ts and now >= _next_nudge_ts:
            _do_nudge()


threading.Thread(target=_controller, daemon=True, name="mouselock-ctrl").start()


# ─── 공개 API ────────────────────────────────────────────────────────
def request_toggle():
    """Alt+1 과 동일한 토글을 코드/HTTP 로 요청(테스트·대체용)."""
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
        s = dict(_state)
    now = time.time()
    remaining = int(s["auto_off_ts"] - now) if (s["active"] and s["auto_off_ts"]) else 0
    return {
        "active": s["active"],
        "supported": s["supported"],
        "nudge_count": s["nudge_count"],
        "remaining_sec": max(0, remaining),
    }
