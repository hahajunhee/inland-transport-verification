"""
자동 F5 입력 서비스
온(on) 상태일 때, 일정 주기(기본 15분)마다 OS 에 F5 키를 직접 입력한다.
- 웹 JS 는 OS 키 입력이 불가하므로, 로컬에서 실행 중인 이 서버가 Windows API
  (ctypes user32.keybd_event)로 현재 활성 창에 F5 를 전송한다.
- 상태는 메모리에만 보관(서버 재시작 시 off 로 초기화) → 의도치 않은 자동 재개 방지.
"""
import os
import threading
import time

_DEFAULT_INTERVAL = 15 * 60  # 15분(초)

_state = {
    "enabled": False,
    "interval_sec": _DEFAULT_INTERVAL,
    "press_count": 0,
    "last_press_ts": None,
    "next_due_ts": None,
    "supported": os.name == "nt",  # Windows 에서만 실제 키 입력 가능
}
_lock = threading.Lock()


def _send_f5_windows():
    """Windows: 활성 창에 F5 키 다운/업 전송 (ctypes)."""
    import ctypes
    VK_F5 = 0x74
    KEYEVENTF_KEYUP = 0x0002
    user32 = ctypes.windll.user32
    user32.keybd_event(VK_F5, 0, 0, 0)                 # key down
    user32.keybd_event(VK_F5, 0, KEYEVENTF_KEYUP, 0)   # key up


def send_f5() -> bool:
    """F5 1회 전송. 성공 시 True. (비-Windows 는 미지원 → False)"""
    if os.name != "nt":
        return False
    try:
        _send_f5_windows()
        return True
    except Exception:
        return False


def _move_mouse_1px_windows():
    """Windows: 커서를 1px 이동했다가 원위치(미세 흔들기 → 활동 감지/절전 방지)."""
    import ctypes
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    pt = wintypes.POINT()
    user32.GetCursorPos(ctypes.byref(pt))
    user32.SetCursorPos(pt.x + 1, pt.y)
    user32.SetCursorPos(pt.x, pt.y)


def move_mouse_1px() -> bool:
    """마우스 1px 이동 1회. 성공 시 True. (비-Windows 는 미지원 → False)"""
    if os.name != "nt":
        return False
    try:
        _move_mouse_1px_windows()
        return True
    except Exception:
        return False


def _loop():
    """백그라운드 루프: enabled 면 interval 마다 F5 입력 + 마우스 1px 이동."""
    while True:
        time.sleep(1)
        with _lock:
            enabled = _state["enabled"]
            interval = _state["interval_sec"]
            next_due = _state["next_due_ts"]
        if not enabled:
            continue
        now = time.time()
        if next_due is None:
            with _lock:
                _state["next_due_ts"] = now + interval
            continue
        if now >= next_due:
            ok1 = send_f5()
            ok2 = move_mouse_1px()
            with _lock:
                if ok1 or ok2:
                    _state["press_count"] += 1
                    _state["last_press_ts"] = now
                _state["next_due_ts"] = now + _state["interval_sec"]


_thread = threading.Thread(target=_loop, daemon=True, name="autopress-f5")
_thread.start()


def _coerce_interval(value):
    """초 단위 정수로 변환(최소 1초). 실패 시 None."""
    try:
        return max(1, int(float(value)))
    except (TypeError, ValueError):
        return None


def set_interval(seconds) -> dict:
    """주기(초) 설정. 켜져 있으면 다음 동작 시점을 새 주기로 다시 잡는다."""
    iv = _coerce_interval(seconds)
    with _lock:
        if iv is not None:
            _state["interval_sec"] = iv
            if _state["enabled"]:
                _state["next_due_ts"] = time.time() + iv
    return get_status()


def set_enabled(enabled: bool, interval_sec=None) -> dict:
    """온/오프 설정. interval_sec 가 오면 함께 적용. 켜면 다음 동작 시점을 지금부터 주기 후로."""
    with _lock:
        iv = _coerce_interval(interval_sec) if interval_sec is not None else None
        if iv is not None:
            _state["interval_sec"] = iv
        _state["enabled"] = bool(enabled)
        if enabled:
            _state["next_due_ts"] = time.time() + _state["interval_sec"]
        else:
            _state["next_due_ts"] = None
    return get_status()


def get_status() -> dict:
    with _lock:
        s = dict(_state)
    return {
        "enabled": s["enabled"],
        "interval_sec": s["interval_sec"],
        "interval_min": round(s["interval_sec"] / 60),
        "press_count": s["press_count"],
        "last_press_ts": s["last_press_ts"],
        "supported": s["supported"],
    }
