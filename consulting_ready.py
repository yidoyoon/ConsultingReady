# -*- coding: utf-8 -*-
"""
ConsultingReady
===========
Excel / PowerPoint 문서를 "저장하고 닫을 때만" 화면 배율(Zoom)을 사용자가 설정한
값으로 고정하거나(또는 '창에 맞춤' 상태로) 저장해 주는 Windows 백그라운드 프로그램.

주요 기능
---------
* 고정 배율을 설정창에서 원하는 값(10~400%, PowerPoint 지원 범위)으로 지정 (기본 100%)
* '창에 맞춤' 토글: 켜면 배율 고정 대신 창에 맞춤 상태로 저장 (배율 고정은 비활성화)
* 트레이 아이콘에 현재 설정(배율 숫자 또는 FIT)을 표시
* 트레이 우클릭 → 설정창 진입 / 일시정지 / 창에 맞춤 토글
* 일시정지 중에는 아무 동작도 하지 않으며 아이콘이 노란색으로 바뀜

동작 방침
---------
사용자가 "의도적으로 저장한 문서를 닫을 때만" 처리합니다. 저장되지 않은 편집
내용을 강제로 저장하거나 버리려는 변경을 저장하지 않습니다.

  * 편집 -> 저장(Ctrl+S) -> 닫기        : 닫는 시점에 적용   (적용)
  * 저장 안 된 변경이 있는 채로 닫기      : 건드리지 않음
  * 열어서 보기만 하고 닫기              : 건드리지 않음
  * 작업 중 Ctrl+S 만 한 경우           : 건드리지 않음 (닫을 때만 적용)

Excel 구현 원리
---------------
닫는 도중에 배율을 저장하면 창이 사라지며 반영되지 않으므로(Excel 특성), BeforeClose
에서 닫기를 잠깐 취소하고 -> 이벤트 밖에서 배율을 바꿔 저장 -> 다시 닫는 방식으로
확실히 반영한다. (Excel 은 '창에 맞춤' 개념이 없어 그 모드에서는 미적용)

PowerPoint 참고
---------------
PowerPoint 는 파일을 열 때마다 배율을 자동 재계산하므로, 특정 배율 고정은 best-effort
이며 다시 열 때 유지가 보장되지 않습니다. '창에 맞춤' 은 PowerPoint 의 기본 동작이라
안정적으로 적용됩니다.
"""

import os
import re
import sys
import gc
import json
import time
import queue
import shutil
import hashlib
import tempfile
import threading
import traceback
import unicodedata
import datetime

# ---------------------------------------------------------------------------
# win32com gen_py 캐시를 쓰기 가능한 임시 폴더로 (frozen exe 대응)
# win32com.client 를 import 하기 전에 설정해야 한다.
# ---------------------------------------------------------------------------
import win32com  # noqa: E402

_GEN_DIR = os.path.join(tempfile.gettempdir(), "consulting_ready_gen_py")
os.makedirs(_GEN_DIR, exist_ok=True)
win32com.__gen_path__ = _GEN_DIR
import win32com.gen_py            # noqa: E402
win32com.gen_py.__path__ = [_GEN_DIR]

import win32com.client            # noqa: E402
import win32com.client.gencache   # noqa: E402
import pythoncom                  # noqa: E402
import win32event                 # noqa: E402
import win32api                   # noqa: E402
import winerror                   # noqa: E402

try:
    win32com.client.gencache.is_readonly = False
except Exception:
    pass

# PowerPoint 가 지원하는 배율 범위 (조사 결과: 10 ~ 400)
ZOOM_MIN = 10
ZOOM_MAX = 400
DEFAULT_ZOOM = 100

# 프로그램 정보
APP_NAME = "ConsultingReady"
APP_VERSION = "1.2"
APP_DESC = ("Excel / PowerPoint 문서를 저장하고 닫을 때\n"
            "화면 배율을 지정한 값으로 고정하고,\n"
            "각 시트의 커서를 A1 으로 이동해 저장합니다.")
APP_AUTHOR = "컨설팅그룹  이도윤 대리"
APP_EMAIL = "ydy@autocrypt.io"

# ---------------------------------------------------------------------------
# 경로 / 설정 / 로그
# ---------------------------------------------------------------------------
_LOCAL = os.environ.get("LOCALAPPDATA", tempfile.gettempdir())
_DATA_DIR = os.path.join(_LOCAL, "ConsultingReady")
os.makedirs(_DATA_DIR, exist_ok=True)
_LOG_PATH = os.path.join(_DATA_DIR, "log.txt")
_CONFIG_PATH = os.path.join(_DATA_DIR, "config.json")

# 이전 이름(OfficeZoomReset)으로 쓰던 설정이 있으면 한 번만 옮겨온다.
try:
    _OLD_CONFIG = os.path.join(_LOCAL, "OfficeZoomReset", "config.json")
    if not os.path.exists(_CONFIG_PATH) and os.path.isfile(_OLD_CONFIG):
        shutil.copyfile(_OLD_CONFIG, _CONFIG_PATH)
except Exception:
    pass


def log(msg):
    line = "[%s] %s" % (
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        msg,
    )
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def _clamp_zoom(z):
    try:
        z = int(round(float(z)))
    except Exception:
        z = DEFAULT_ZOOM
    return max(ZOOM_MIN, min(ZOOM_MAX, z))


# ---------------------------------------------------------------------------
# 런타임 상태
#   _target_zoom / _fit_mode / _paused 는 감시(STA) 스레드에서 읽고 트레이
#   스레드에서 쓴다. (int/bool 대입은 GIL 하에서 원자적이라 별도 락 불필요)
# ---------------------------------------------------------------------------
_target_zoom = DEFAULT_ZOOM
_fit_mode = False
_paused = False
_toast_enabled = True
# 파일명 키워드 필터
_filter_enabled = False
_filter_mode = "exclude"     # "exclude"=키워드가 들어간 파일 제외 / "include"=들어간 파일만 처리
_filter_keywords = ""        # 사용자가 입력한 그대로 (예: "구글, 제미나이")
_filter_list = []            # 비교용으로 정규화한 키워드 목록
# 기울어진 도형(딱지) 검사
_stamp_check = True
# PowerPoint 메모 창 접고 저장 (발표 준비 상태)
_hide_notes = True
# 수정 전 원본 백업
_backup_enabled = False
_backup_dir = ""          # 비어 있으면 %LOCALAPPDATA%\ConsultingReady\backup
_backup_keep = 20         # 원본 1개당 보관할 최근 백업 개수
_icon = None
_settings_open = False
_about_open = False
_mutex = None

# 문서 상태 (모두 감시 스레드에서만 접근)
_session_saved = set()   # 이번 세션에 사용자가 '저장'한 문서
_session_edited = set()  # 이번 세션에 '편집'된 문서 (AutoSave 문서 판단용)
_pending = []            # [(wb, key)] Excel: 닫기 취소됨 -> 처리 대기
_allow_close = set()
_tried = set()


def load_config():
    z, fit, toast = DEFAULT_ZOOM, False, True
    fen, fmode, fpat, stamp, notes = False, "exclude", "", True, True
    ben, bdir, bkeep = False, "", 20
    try:
        with open(_CONFIG_PATH, encoding="utf-8-sig") as f:   # BOM 있어도 처리
            d = json.load(f)
        z = _clamp_zoom(d.get("zoom", DEFAULT_ZOOM))
        fit = bool(d.get("fit_to_window", False))
        toast = bool(d.get("toast", True))
        fen = bool(d.get("filter_enabled", False))
        fmode = d.get("filter_mode", "exclude")
        if fmode not in ("exclude", "include"):
            fmode = "exclude"
        fkw = d.get("filter_keywords")
        if fkw is None:
            # 1.2 이하에서 저장한 정규식 필터 값을 키워드로 옮긴다.
            fkw = keywords_from_regex(d.get("filter_pattern", ""))
        fpat = str(fkw or "")
        stamp = bool(d.get("stamp_check", True))
        notes = bool(d.get("hide_notes", True))
        ben = bool(d.get("backup_enabled", False))
        bdir = str(d.get("backup_dir", "") or "")
        try:
            bkeep = max(1, min(200, int(d.get("backup_keep", 20))))
        except Exception:
            bkeep = 20
    except Exception:
        pass
    return z, fit, toast, fen, fmode, fpat, stamp, notes, ben, bdir, bkeep


def save_config():
    try:
        with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"zoom": int(_target_zoom),
                       "fit_to_window": bool(_fit_mode),
                       "toast": bool(_toast_enabled),
                       "filter_enabled": bool(_filter_enabled),
                       "filter_mode": _filter_mode,
                       "filter_keywords": _filter_keywords,
                       "stamp_check": bool(_stamp_check),
                       "hide_notes": bool(_hide_notes),
                       "backup_enabled": bool(_backup_enabled),
                       "backup_dir": _backup_dir,
                       "backup_keep": int(_backup_keep)},
                      f, ensure_ascii=False)
    except Exception:
        pass


_KW_SPLIT = re.compile(r"[,，、;\n]+")   # 쉼표(전각 포함), 세미콜론, 줄바꿈


def _norm_name(s):
    """비교용 정규화: 한글 자모 결합(NFC) · 대소문자 무시 · 공백 전부 제거.
    그래서 '통합 문서1.xlsx' 와 키워드 '통합문서' 처럼 띄어쓰기만 달라도 일치한다."""
    s = unicodedata.normalize("NFC", str(s or ""))
    return "".join(s.casefold().split())


def parse_keywords(text):
    """'구글, 제미나이' → ['구글', '제미나이'] (정규화 · 빈 값과 중복 제외)."""
    out = []
    for part in _KW_SPLIT.split(str(text or "")):
        k = _norm_name(part)
        if k and k not in out:
            out.append(k)
    return out


def keywords_from_regex(pattern):
    """1.2 이하의 정규식 필터 값을 키워드 문자열로 옮긴다.
    '최종|배포' → '최종, 배포' / '^통합문서1\\.xlsx$' → '통합문서1.xlsx'
    와일드카드 · 묶음 · 문자 클래스처럼 글자로 옮길 수 없는 조각은 버린다."""
    words = []
    for part in str(pattern or "").split("|"):
        part = part.strip()
        if part.startswith("^"):
            part = part[1:]
        if part.endswith("$") and not part.endswith("\\$"):
            part = part[:-1]
        word, ok, i = [], True, 0
        while i < len(part):
            ch = part[i]
            if ch == "\\":
                nxt = part[i + 1:i + 2]
                if not nxt or nxt.isalnum():     # \\d, \\w 같은 기호는 글자로 옮길 수 없다
                    ok = False
                    break
                word.append(nxt)
                i += 2
                continue
            if ch in ".*+?()[]{}^$":
                ok = False
                break
            word.append(ch)
            i += 1
        w = "".join(word).strip()
        if ok and w:
            words.append(w)
    return ", ".join(words)


def refresh_filter():
    """입력된 키워드 문자열을 비교용 목록으로 갱신한다."""
    global _filter_list
    _filter_list = parse_keywords(_filter_keywords)


def _filter_desc():
    """로그용 필터 설명."""
    if not _filter_enabled:
        return "꺼짐"
    return "%s [%s]" % ("제외" if _filter_mode == "exclude" else "허용",
                        _filter_keywords.strip() or "키워드 없음")


def filter_skip_reason(name):
    """파일명 키워드 필터 때문에 건너뛰어야 하면 사유, 아니면 None.
    키워드 중 하나라도 파일명에 들어 있으면 '일치'로 본다."""
    if not _filter_enabled or not _filter_list:
        return None
    n = _norm_name(name)
    hit = next((k for k in _filter_list if k in n), None)
    if _filter_mode == "include":
        return None if hit else "허용 키워드 없음"
    return ("제외 키워드 '%s' 포함" % hit) if hit else None


def _safe_name(doc):
    try:
        return doc.Name
    except Exception:
        return "?"


# ---------------------------------------------------------------------------
# 수정 전 원본 백업
#   문서를 우리가 손대기(apply_excel / apply_ppt) 전에 원본을 그대로 복사해 둔다.
#   설계 원칙:
#   * 동기 복사 + 원자적 교체(.part → os.replace). 부분 파일이 최종 이름을 얻지 못한다.
#   * 메타데이터(mtime)는 복사하지 않는다. 백업 파일의 mtime = '백업한 시각' 이어야
#     보관 정리가 올바르게 동작한다.
#   * 삭제는 우리가 만든 이름(_ozr_ 접두사)만 대상으로 한다(화이트리스트).
#   * 실패하면 문서를 수정하지 않는다(fail-closed).
# ---------------------------------------------------------------------------
_BACKUP_PREFIX = "_ozr_"
_BACKUP_INDEX = "_ozr_index.txt"
_BACKUP_MAX_BYTES = 500 * 1024 * 1024      # 500MB 초과는 백업하지 않음
_BACKUP_RE = re.compile(
    r"^_ozr_.+_[0-9a-f]{8}_\d{8}_\d{6}(?:_\d{2})?\.[^.]+$", re.IGNORECASE)


def _is_backup_name(name):
    try:
        return bool(_BACKUP_RE.match(name or ""))
    except Exception:
        return False


def backup_root():
    return _backup_dir.strip() or os.path.join(_DATA_DIR, "backup")


def _is_remote_path(path):
    """SharePoint/OneDrive 등 URL 로 열린 문서인가(로컬 복사 불가)."""
    p = (path or "").lower()
    return (p.startswith("http://") or p.startswith("https://")
            or "@ssl" in p or "davwwwroot" in p)


def _safe_stem(s):
    s = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", s or "")
    s = s.rstrip(". ")
    return s[:80] or "doc"


def _append_backup_index(root, backup_name, src):
    try:
        with open(os.path.join(root, _BACKUP_INDEX), "a", encoding="utf-8") as f:
            f.write("%s\t%s\t%s\n"
                    % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       backup_name, src))
    except Exception:
        pass


def _cleanup_backups(root, tag):
    """같은 원본(tag)의 백업을 최신 _backup_keep 개만 남긴다.
    우리가 만든 이름만 지우며, 가장 최신 백업은 절대 지우지 않는다."""
    keep = max(1, int(_backup_keep))
    try:
        items = []
        for nm in os.listdir(root):
            if not _is_backup_name(nm) or ("_%s_" % tag) not in nm:
                continue
            p = os.path.join(root, nm)
            try:
                if os.path.isfile(p):
                    items.append((os.path.getmtime(p), p))
            except Exception:
                pass
        items.sort(key=lambda t: t[0], reverse=True)     # 최신이 앞
        for _mt, p in items[keep:]:                      # keep>=1 → 최신은 항상 보존
            try:
                os.remove(p)
            except Exception:
                pass
    except Exception:
        pass


def backup_document(path):
    """수정 전 원본을 백업한다.
    반환 (state, message, backup_path):
      'skip' = 백업할 필요 없음(정상)  /  'ok' = 백업함  /  'fail' = 실패"""
    if not _backup_enabled:
        return "skip", "백업 꺼짐", None
    if _is_remote_path(path):
        return "skip", "온라인 문서(서버 버전 기록 사용)", None
    try:
        src = os.path.abspath(path)
        if not os.path.isfile(src):
            return "fail", "원본 파일을 찾을 수 없음", None
        size = os.path.getsize(src)
        if size > _BACKUP_MAX_BYTES:
            return "fail", "파일이 너무 큼(%.0fMB)" % (size / 1048576.0), None

        root = backup_root()
        try:
            os.makedirs(root, exist_ok=True)
        except Exception as e:
            return "fail", "백업 폴더를 만들 수 없음: %s" % e, None
        if os.path.normcase(os.path.abspath(root)) == \
           os.path.normcase(os.path.dirname(src)):
            return "fail", "백업 폴더가 원본 폴더와 같음", None

        base = os.path.basename(src)
        stem, ext = os.path.splitext(base)
        # 원본 경로를 짧게 구분하기 위한 지문(보안 용도 아님).
        # 같은 이름의 다른 폴더 파일이 섞이지 않게 하는 그룹 키로만 쓴다.
        tag = hashlib.blake2b(
            os.path.normcase(src).encode("utf-16le"), digest_size=4).hexdigest()
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

        dst = None
        for n in range(0, 20):
            cand = os.path.join(root, "%s%s_%s_%s%s%s" % (
                _BACKUP_PREFIX, _safe_stem(stem), tag, ts,
                "" if n == 0 else "_%02d" % n, ext))
            if not os.path.exists(cand):
                dst = cand
                break
        if dst is None:
            return "fail", "백업 이름을 만들지 못함", None

        tmp = dst + ".part"
        try:
            shutil.copyfile(src, tmp)        # mtime 은 복사하지 않는다(의도적)
            if os.path.getsize(tmp) != size:
                raise IOError("복사 크기 불일치")
            os.replace(tmp, dst)
        except Exception:
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except Exception:
                pass
            raise

        _append_backup_index(root, os.path.basename(dst), src)
        _cleanup_backups(root, tag)
        return "ok", os.path.basename(dst), dst
    except Exception as e:
        return "fail", str(e), None


def discard_backup(bpath, why):
    """수정할 게 없어 저장하지 않은 경우, 방금 만든 백업을 되돌린다."""
    if not bpath:
        return
    try:
        os.remove(bpath)
        log("백업 되돌림(%s): %s" % (why, os.path.basename(bpath)))
    except Exception:
        pass


def notify_backup_fail(kind, name, why):
    log("%s: 백업 실패 → 수정하지 않고 닫음 (%s) - %s" % (kind, name, why))
    show_toast("백업에 실패해 파일을 수정하지 않았습니다.",
               "%s\n%s\n설정에서 백업 폴더를 확인해 주세요." % (name, why),
               warn=True, force=True)


def _autosave_on(doc):
    """OneDrive / SharePoint 문서의 '자동 저장' 켜짐 여부.
    (해당 속성이 없는 예전 버전/로컬 파일은 False)"""
    try:
        return bool(doc.AutoSaveOn)
    except Exception:
        return False


def _forget_doc(k):
    for s in (_session_saved, _session_edited):
        s.discard(k)


# ---------------------------------------------------------------------------
# 토스트 알림 (화면 우측 하단에 3초간 표시)
#   COM 감시 스레드를 막지 않도록 전용 스레드에서 하나씩 순서대로 표시한다.
# ---------------------------------------------------------------------------
TOAST_HOLD_MS = 3000          # 표시 유지 시간(3초)
_toast_q = queue.Queue()


def show_toast(title, detail, warn=False, force=False):
    """토스트 표시 요청(설정에서 꺼져 있으면 무시).
    warn=True 면 주의(주황) 색으로 표시하고 조금 더 오래 띄운다.
    force=True 는 알림을 꺼둔 경우에도 표시한다(백업 실패 같은 오류 전용).
    오류를 조용히 삼키면 기능이 멈춘 걸 알 수 없기 때문이다."""
    if not _toast_enabled and not force:
        return
    try:
        _toast_q.put_nowait((title, detail, warn))
    except Exception:
        pass


def _toast_loop(stop_event):
    while not stop_event.is_set():
        try:
            item = _toast_q.get(timeout=0.5)
        except Exception:
            continue
        if item is None:
            break
        try:
            _show_toast_window(item[0], item[1],
                               item[2] if len(item) > 2 else False)
        except Exception:
            log("토스트 표시 실패\n" + traceback.format_exc())


def _work_area():
    """작업 표시줄을 제외한 화면 영역 (left, top, right, bottom)."""
    try:
        mon = win32api.MonitorFromPoint((0, 0), 1)   # MONITOR_DEFAULTTOPRIMARY
        return win32api.GetMonitorInfo(mon)["Work"]
    except Exception:
        return None


def _show_toast_window(title, detail, warn=False):
    import tkinter as tk

    W, H, MARGIN = 340, 92, 16
    root = tk.Tk()
    root.withdraw()
    root.overrideredirect(True)          # 제목 표시줄 없음
    try:
        root.attributes("-topmost", True)
        root.attributes("-alpha", 0.0)   # 페이드 인 시작
    except Exception:
        pass

    area = _work_area()
    if area:
        x = area[2] - W - MARGIN
        y = area[3] - H - MARGIN
    else:
        x = root.winfo_screenwidth() - W - MARGIN
        y = root.winfo_screenheight() - H - 60
    root.geometry("%dx%d+%d+%d" % (W, H, x, y))

    if warn:
        BG, FG, SUB, ACCENT = "#2e2617", "#ffffff", "#d8c9a8", "#e5a50a"
        mark = "⚠  "
    else:
        BG, FG, SUB, ACCENT = "#202b23", "#ffffff", "#b9c7bd", "#21bf6a"
        mark = "✔  "
    outer = tk.Frame(root, bg=ACCENT)
    outer.pack(fill="both", expand=True)
    card = tk.Frame(outer, bg=BG)
    card.pack(fill="both", expand=True, padx=(4, 0))       # 왼쪽 액센트 바

    tk.Label(card, text=mark + title, bg=BG, fg=FG,
             font=("Malgun Gothic", 11, "bold"), anchor="w",
             justify="left").pack(fill="x", padx=14, pady=(14, 2))
    tk.Label(card, text=detail, bg=BG, fg=SUB,
             font=("Malgun Gothic", 9), anchor="w", justify="left",
             wraplength=W - 34).pack(fill="x", padx=14, pady=(0, 12))

    # 클릭하면 즉시 닫기
    def close(_evt=None):
        try:
            root.destroy()
        except Exception:
            pass
    for w in (root, outer, card):
        w.bind("<Button-1>", close)

    root.deiconify()

    def fade(value, step, after_done):
        try:
            root.attributes("-alpha", value)
        except Exception:
            pass
        nxt = value + step
        if (step > 0 and nxt < 0.96) or (step < 0 and nxt > 0.05):
            root.after(15, fade, nxt, step, after_done)
        else:
            after_done()

    def hold():
        root.after(TOAST_HOLD_MS * (2 if warn else 1),
                   lambda: fade(0.96, -0.08, close))

    root.after(0, fade, 0.0, 0.12, hold)
    try:
        root.mainloop()
    finally:
        # Tk 객체는 반드시 '만든 스레드'에서 정리해야 한다.
        # (다른 스레드의 gc.collect() 가 수거하면 Tcl_AsyncDelete 로 죽을 수 있음)
        try:
            root.destroy()
        except Exception:
            pass
        close = fade = hold = None
        card = outer = None
        root = None
        gc.collect()


# ---------------------------------------------------------------------------
# Excel  ('창에 맞춤' 모드에서는 미적용 - Excel 에는 해당 개념이 없음)
# ---------------------------------------------------------------------------
def excel_key(wb):
    try:
        if not wb.Path:
            return None
        if wb.ReadOnly:
            return None
        if _is_backup_name(wb.Name):     # 백업본은 대상에서 제외(재귀 방지)
            return None
        return wb.FullName
    except Exception:
        return None


def apply_excel(wb):
    """모든 시트에 배율(창에 맞춤 모드가 아닐 때)과 A1 셀 선택을 적용한다.
    반환: (무언가 바뀌었나, 배율 적용함, A1 적용함)"""
    target = _target_zoom
    do_zoom = not _fit_mode          # 창에 맞춤 모드면 배율은 건드리지 않고 A1 만
    changed = zoom_done = a1_done = False
    try:
        original = None
        try:
            original = wb.Application.ActiveSheet
        except Exception:
            original = None
        try:
            win = wb.Windows(1)
        except Exception:
            win = None
        if win is None:
            return (False, False, False)

        try:
            sheets = list(wb.Worksheets)
        except Exception:
            sheets = []

        for ws in sheets:
            try:
                ws.Activate()
            except Exception:
                continue                 # 숨김 시트 등은 건너뜀

            if do_zoom:
                try:
                    if win.Zoom != target:
                        win.Zoom = target
                        changed = True
                    zoom_done = True
                except Exception:
                    pass                 # 차트 시트 등 배율 불가

            # 커서를 A1 로
            try:
                addr = None
                try:
                    addr = wb.Application.ActiveCell.Address
                except Exception:
                    addr = None
                if addr != "$A$1":
                    ws.Range("A1").Select()
                    changed = True
                a1_done = True
            except Exception:
                pass                     # 보호된 시트 등에서 선택 불가

            # 화면도 좌상단으로
            try:
                if win.ScrollRow != 1:
                    win.ScrollRow = 1
                    changed = True
                if win.ScrollColumn != 1:
                    win.ScrollColumn = 1
                    changed = True
            except Exception:
                pass

        # 마지막으로 첫 번째 시트를 활성화 (다음에 열면 1번 시트가 보이도록)
        try:
            cur_name = None
            try:
                cur = wb.Application.ActiveSheet
                cur_name = cur.Name if cur is not None else None
            except Exception:
                cur_name = None
            first = None
            for getter in (lambda: wb.Sheets(1), lambda: wb.Worksheets(1)):
                try:
                    cand = getter()
                    cand.Activate()          # 숨김 시트면 예외 → 다음 후보
                    first = cand
                    break
                except Exception:
                    continue
            if first is not None:
                try:
                    if cur_name is not None and first.Name != cur_name:
                        changed = True
                except Exception:
                    pass
            elif original is not None:
                original.Activate()          # 첫 시트 활성화 실패 시 원래대로
        except Exception:
            pass
    except Exception:
        log("Excel: 배율/A1 적용 중 예외\n" + traceback.format_exc())
    return (changed, zoom_done, a1_done)


class ExcelEvents:
    def OnWorkbookBeforeSave(self, Wb, SaveAsUI, Cancel):
        k = excel_key(Wb)
        if k:
            _session_saved.add(k)

    def OnSheetChange(self, Sh, Target):
        # 셀이 바뀔 때마다 호출된다. AutoSave(OneDrive/SharePoint) 문서는
        # 사용자가 Ctrl+S 를 누르지 않아 저장 이벤트가 잡히지 않으므로,
        # '이 문서를 편집했다'는 사실을 여기서 기록해 둔다.
        try:
            _session_edited.add(Sh.Parent.FullName)
        except Exception:
            pass

    def OnWorkbookBeforeClose(self, Wb, Cancel):
        if _paused:
            return                       # 일시정지: 아무것도 안 함
        k = excel_key(Wb)
        if not k:
            return
        if k in _allow_close:
            _allow_close.discard(k)
            _forget_doc(k)
            return

        name = _safe_name(Wb)
        skip = filter_skip_reason(name)
        if skip:
            log("Excel: 파일명 필터(%s) → 건드리지 않음 (%s)" % (skip, name))
            _forget_doc(k)
            return

        autosave = _autosave_on(Wb)
        try:
            clean = bool(Wb.Saved)
        except Exception:
            clean = False

        if autosave:
            # 자동 저장 문서: 이미 계속 저장되고 있으므로 '저장 안 된 변경'이라는
            # 개념이 없다. 사용자가 편집(또는 저장)한 문서만 대상으로 한다.
            if not (k in _session_saved or k in _session_edited):
                log("Excel: [AutoSave] 편집 없음(조회만) → 건드리지 않음 (%s)" % name)
                _forget_doc(k)
                return
        else:
            if not clean:
                log("Excel: 저장 안 된 변경 있음 → 건드리지 않음 (%s)" % name)
                _forget_doc(k)
                return
            if k not in _session_saved:
                log("Excel: 이번 세션 저장 없음(조회만) → 건드리지 않음 (%s)" % name)
                _forget_doc(k)
                return

        if k in _tried:
            log("Excel: 재시도 방지 → 그대로 닫음 (%s)" % name)
            _tried.discard(k)
            _forget_doc(k)
            return

        # 배율은 활성 시트만 보고 판단할 수 있지만 A1 은 시트마다 달라서
        # 여기서 미리 알 수 없다. 일단 닫기를 취소하고 실제 처리에서 확인한다.
        # (바뀐 게 없으면 저장하지 않고 그대로 닫으므로 불필요한 저장은 없다)
        if all(existing is not Wb for existing, _kk in _pending):
            _pending.append((Wb, k))
        log("Excel: 닫기 취소 → 배율/A1 적용 예약%s (%s)"
            % (" [AutoSave]" if autosave else "", name))
        return True                      # 이번 닫기는 취소


def _open_count(app, key):
    """사용자에게 '열려 있는 문서' 수.
    Excel 은 숨은 워크북(PERSONAL.XLSB 등)을 제외하기 위해 '보이는 창' 수를 센다.
    알 수 없으면 -1 (죽는 중일 수 있음)."""
    try:
        if key == "Excel":
            wins = app.Windows
            n = 0
            cnt = wins.Count
            for i in range(1, cnt + 1):
                try:
                    if wins.Item(i).Visible:
                        n += 1
                except Exception:
                    pass
            return n
        else:
            return app.Presentations.Count
    except Exception:
        return -1


def _release_app(entry, quit_first=False):
    """앱 이벤트 sink 연결을 끊고 참조를 놓아 프로세스가 종료될 수 있게 한다.
    quit_first=True 면 앱을 먼저 Quit 한다(마지막 문서를 닫는 경우)."""
    app = entry.get("app")
    sink = entry.get("sink")
    if quit_first and app is not None:
        try:
            app.Quit()
        except Exception as e:
            log("앱 종료(Quit) 오류: %s" % e)
    if sink is not None:
        try:
            sink.close()      # 이벤트 연결 해제(unadvise) - 반드시 필요
        except Exception:
            pass
    entry["sink"] = None
    entry["app"] = None
    app = None
    sink = None
    # COM 개체가 실제로 해제되도록 gc + 메시지 펌프
    for _ in range(3):
        gc.collect()
        try:
            pythoncom.PumpWaitingMessages()
        except Exception:
            pass


def process_pending(excel_entry):
    while _pending:
        wb, k = _pending.pop(0)
        _tried.add(k)
        name = _safe_name(wb)
        try:
            app = None
            last = False
            try:
                app = wb.Application
                last = (_open_count(app, "Excel") <= 1)
            except Exception:
                app = None
                last = False

            # 우리가 손대기 전에 원본을 백업한다(디스크는 아직 원본 상태).
            bstate, bwhy, bpath = backup_document(k)
            if bstate == "fail":
                # 백업 없이는 수정하지 않는다. 닫기만 정상적으로 완수한다.
                notify_backup_fail("Excel", name, bwhy)
            else:
                if bstate == "ok":
                    log("Excel: 원본 백업 (%s)" % bwhy)
                changed, zoom_done, a1_done = apply_excel(wb)
                if changed:
                    try:
                        wb.Saved = False
                    except Exception:
                        pass
                    wb.Save()

                    parts = []
                    if zoom_done:
                        parts.append("배율 %d%%" % _target_zoom)
                    if a1_done:
                        parts.append("A1 셀 · 첫 시트")
                    summary = " · ".join(parts) if parts else "설정 적용"
                    show_toast("저장 완료",
                               "%s\n%s 후 저장했습니다." % (name, summary))
                else:
                    log("Excel: 이미 설정대로임 → 저장 없이 닫음 (%s)" % name)
                    discard_backup(bpath, "변경 없음")

            if last and app is not None and excel_entry.get("app") is not None:
                # 마지막(보이는) 워크북: 저장 후 Excel 을 완전히 종료한다.
                # (닫기를 취소했었기 때문에 우리가 Quit 하지 않으면 Excel 이 빈 채로 남는다)
                app = None
                wb = None
                _release_app(excel_entry, quit_first=True)
                # 종료 중인 인스턴스를 곧바로 다시 잡지 않도록 잠깐 쉰다
                excel_entry["next"] = time.time() + 3.0
                _tried.discard(k)
                _forget_doc(k)
                log("Excel: 마지막 문서 처리 후 Excel 종료 (%s)" % name)
            else:
                _allow_close.add(k)
                wb.Close()
                _tried.discard(k)
                log("Excel: 배율/A1 처리 후 닫음 (%s)" % name)
        except Exception as e:
            # 실패해도 상태를 남기지 않는다. (남으면 그 문서가 계속 건너뛰어짐)
            _allow_close.discard(k)
            _tried.discard(k)
            _forget_doc(k)
            log("Excel: 처리 실패 → 다음 닫기 허용 (%s) %s" % (name, e))


# ---------------------------------------------------------------------------
# PowerPoint (best-effort)
# ---------------------------------------------------------------------------
def ppt_key(pres):
    try:
        if not pres.Path:
            return None
        if pres.ReadOnly:
            return None
        if _is_backup_name(pres.Name):   # 백업본은 대상에서 제외(재귀 방지)
            return None
        return pres.FullName
    except Exception:
        return None


PP_VIEW_NORMAL = 9        # ppViewNormal


def _notes_split(pres):
    """보통 보기에서 슬라이드 영역이 창 높이에서 차지하는 비율(%).
    100 이면 메모 창이 접힌 상태다. 보통 보기가 아니거나 읽을 수 없으면 None."""
    try:
        w = pres.Windows(1)
        if w.ViewType != PP_VIEW_NORMAL:
            return None
        return int(w.SplitVertical)
    except Exception:
        return None


def hide_notes_pane(pres):
    """메모(발표자 노트) 창을 접어 발표 준비 상태로 만든다. 접었으면 True.

    창의 SplitVertical(슬라이드 영역 비율)을 100 으로 지정해 접는다. 이 상태는
    저장할 때 viewProps.xml 의 horzBarState="maximized" 로 기록되어 다시 열어도
    유지된다. (뷰 변경만으로는 문서가 '수정됨'이 되지 않으므로 저장 직전에
     Saved=False 를 함께 지정해야 실제로 기록된다.)

    리본의 '메모' 버튼(ExecuteMso "ShowNotes")으로 접으면 OneDrive/SharePoint
    (자동 저장) 문서에서는 닫기 이벤트 안에서 화면만 바뀌고 파일에는 기록되지
    않았다. 그래서 문서 창 속성을 직접 바꾼다. 보통 보기가 아니어서
    SplitVertical 을 쓸 수 없을 때만 리본 방식으로 접는다."""
    if not _hide_notes:
        return False
    try:
        split = _notes_split(pres)
        if split is not None:
            if split >= 100:
                return False
            pres.Windows(1).SplitVertical = 100
            return True
        if _notes_pane_open(pres) is True:   # 창 활성화 포함
            pres.Application.CommandBars.ExecuteMso("ShowNotes")
            return True
    except Exception:
        log("PPT: 메모창 접기 실패\n" + traceback.format_exc())
    return False


def _notes_pane_open(pres):
    """메모 창이 펼쳐져 있으면 True (판단 불가면 None)."""
    split = _notes_split(pres)
    if split is not None:
        return split < 100
    try:
        app = pres.Application
        try:
            pres.Windows(1).Activate()   # 대상 창 기준으로 상태를 읽는다
        except Exception:
            pass
        return bool(app.CommandBars.GetPressedMso("ShowNotes"))
    except Exception:
        return None


def warn_notes_open(pres, name):
    """저장하지 않은 문서용: 파일은 건드리지 않고 메모 창이 열려 있으면 알림만."""
    if not _hide_notes:
        return
    if _notes_pane_open(pres) is True:
        log("PPT: 메모 창 열려 있음(저장하지 않음) - %s" % name)
        show_toast("메모 창이 열려 있습니다.",
                   "%s\n발표 준비 상태로 만들려면 메모 창을 접고 저장하세요."
                   % name, warn=True)


def apply_ppt(pres):
    """현재 모드('창에 맞춤' 또는 배율)에 맞게 PPT 를 설정. 바뀌었으면 True."""
    changed = hide_notes_pane(pres)
    try:
        windows = list(pres.Windows)
    except Exception:
        windows = []
    for w in windows:
        try:
            view = w.View
            if _fit_mode:
                if not view.ZoomToFit:          # 아직 창맞춤이 아니면
                    view.ZoomToFit = True
                    changed = True
            else:
                try:
                    if view.ZoomToFit:
                        view.ZoomToFit = False
                        changed = True
                except Exception:
                    pass
                if view.Zoom != _target_zoom:
                    view.Zoom = _target_zoom
                    changed = True
        except Exception:
            pass
    return changed


MSO_GROUP = 6            # msoGroup
_ROT_EPS = 0.5           # 이 각도 미만은 회전 없음으로 본다


def _scan_tilted(shapes, where, found, depth=0):
    """도형 목록을 훑어 회전된(기울어진) 도형을 found 에 모은다.
    그룹 안쪽도 확인한다."""
    try:
        n = shapes.Count
    except Exception:
        return
    for i in range(1, n + 1):
        try:
            sh = shapes.Item(i)
        except Exception:
            continue
        try:
            rot = float(sh.Rotation or 0.0) % 360.0
        except Exception:
            rot = 0.0
        if min(rot, 360.0 - rot) >= _ROT_EPS:
            try:
                nm = sh.Name
            except Exception:
                nm = "?"
            found.append((where, nm, rot))
        # 그룹 내부도 검사
        if depth < 3:
            try:
                if sh.Type == MSO_GROUP:
                    _scan_tilted(sh.GroupItems, where, found, depth + 1)
            except Exception:
                pass


def find_tilted_shapes(pres, limit=200):
    """프레젠테이션의 모든 슬라이드에서 기울어진 도형을 찾는다.
    반환: [(슬라이드번호, 도형이름, 각도), ...]"""
    found = []
    try:
        slides = pres.Slides
        cnt = slides.Count
    except Exception:
        return found
    for i in range(1, cnt + 1):
        if len(found) >= limit:
            break
        try:
            sl = slides.Item(i)
            _scan_tilted(sl.Shapes, i, found)
        except Exception:
            continue
    return found


def check_stamps(pres, name):
    """기울어진 도형이 있으면 경고 토스트를 띄운다."""
    if not _stamp_check:
        return
    try:
        found = find_tilted_shapes(pres)
    except Exception:
        log("PPT: 딱지 검사 중 예외\n" + traceback.format_exc())
        return
    if not found:
        log("PPT: 기울어진 도형 없음 (%s)" % name)
        return
    spots = []
    for slide_no, shname, rot in found[:3]:
        spots.append("슬라이드 %d '%s' (%.0f°)" % (slide_no, shname, rot))
    more = "" if len(found) <= 3 else " 외 %d개" % (len(found) - 3)
    log("PPT: 기울어진 도형 %d개 발견 (%s) - %s%s"
        % (len(found), name, " / ".join(spots), more))
    show_toast("기울어진 도형(딱지)가 남아있습니다.",
               "%s\n%s%s" % (name, " / ".join(spots), more),
               warn=True)


class PowerPointEvents:
    def OnPresentationBeforeSave(self, Pres, Cancel):
        k = ppt_key(Pres)
        if k:
            _session_saved.add(k)

    def OnPresentationBeforeClose(self, Pres, Cancel):
        if _paused:
            return
        k = ppt_key(Pres)
        if not k:
            return
        name = _safe_name(Pres)
        skip = filter_skip_reason(name)
        if skip:
            log("PPT: 파일명 필터(%s) → 건드리지 않음 (%s)" % (skip, name))
            _forget_doc(k)
            return

        try:
            clean = (int(Pres.Saved) != 0)
        except Exception:
            clean = False
        mode = "창에 맞춤" if _fit_mode else ("배율 %d%%" % _target_zoom)
        try:
            if not (clean and k in _session_saved):
                # 저장하지 않은 문서: 파일은 전혀 건드리지 않고 검사만 한다.
                log("PPT: %s → 파일 수정 없이 검사만 (%s)"
                    % ("저장 안 된 변경 있음" if not clean
                       else "이번 세션 저장 없음(조회만)", name))
                warn_notes_open(Pres, name)
            else:
                # 메모창/배율을 손대기 전에 원본을 백업한다.
                bstate, bwhy, bpath = backup_document(k)
                if bstate == "fail":
                    # 백업 없이는 수정하지 않는다. apply_ppt 를 호출하지 않으므로
                    # 문서가 '수정됨'이 되지 않아 저장 프롬프트도 뜨지 않는다.
                    notify_backup_fail("PPT", name, bwhy)
                elif apply_ppt(Pres):
                    if bstate == "ok":
                        log("PPT: 원본 백업 (%s)" % bwhy)
                    try:
                        # 뷰 변경만으로는 '수정됨'이 되지 않아 Save() 가 무시된다.
                        # 강제로 표시해야 메모창/배율 상태가 파일에 기록된다.
                        try:
                            Pres.Saved = False
                        except Exception:
                            pass
                        Pres.Save()
                        log("PPT: (best-effort) %s%s 저장 (%s)"
                            % (mode, " · 메모창 접음" if _hide_notes else "", name))
                        show_toast("저장 완료",
                                   "%s\n%s 적용 후 저장했습니다." % (name, mode))
                    except Exception as e:
                        # 저장에 실패해도 백업은 남겨 둔다(가장 필요한 순간이다).
                        log("PPT: 저장 실패 (%s) %s" % (name, e))
                else:
                    log("PPT: 이미 %s → 변경 없음 (%s)" % (mode, name))
                    discard_backup(bpath, "변경 없음")
            # 딱지(기울어진 도형) 검사는 수정 여부와 관계없이 항상 수행한다.
            # (파일을 읽기만 하므로 문서를 건드리지 않는다)
            check_stamps(Pres, name)
        finally:
            _forget_doc(k)


# ---------------------------------------------------------------------------
# 감시 루프
# ---------------------------------------------------------------------------
_APPS = [
    ("Excel", "Excel.Application", ExcelEvents),
    ("PowerPoint", "PowerPoint.Application", PowerPointEvents),
]
_RETRY_DELAY = 5.0


def _is_alive(app):
    try:
        _ = app.Name
        return True
    except Exception:
        return False


def _is_visible(app):
    try:
        return bool(app.Visible)
    except Exception:
        return False


_DOC_EXT = {
    "Excel": (".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".xltm", ".csv"),
    "PowerPoint": (".ppt", ".pptx", ".pptm", ".potx", ".ppsx", ".ppsm"),
}


def _find_app_via_rot(key):
    """열려 있는 문서를 통해 실제 앱 인스턴스를 찾는다.

    GetActiveObject 는 ROT(실행 중 개체 테이블)의 '첫' 인스턴스만 돌려준다.
    종료 중이거나 문서가 없는 낡은 인스턴스가 ROT 에 남아 있으면 계속 그것만
    반환되어, 사용자가 새로 연 Excel 을 영영 감시하지 못하게 된다.
    그래서 문서(.xlsx 등) 모니커를 직접 훑어 살아있는 인스턴스를 찾는다."""
    exts = _DOC_EXT.get(key, ())
    if not exts:
        return None
    try:
        rot = pythoncom.GetRunningObjectTable()
        ctx = pythoncom.CreateBindCtx(0)
        try:
            monikers = list(rot.EnumRunning())
        except Exception:
            monikers = list(rot)
    except Exception:
        return None

    for mk in monikers:
        try:
            name = mk.GetDisplayName(ctx, None)
        except Exception:
            continue
        if not name or not name.lower().endswith(exts):
            continue
        try:
            unk = rot.GetObject(mk)
            # ROT 는 IUnknown 을 주므로 IDispatch 로 변환해야 한다
            doc = win32com.client.Dispatch(
                unk.QueryInterface(pythoncom.IID_IDispatch))
            app = doc.Application
            doc = None
            unk = None
        except Exception:
            continue
        try:
            if _is_alive(app) and _should_hold(app, key):
                return app
        except Exception:
            pass
    return None


def _should_hold(app, key):
    """이 인스턴스를 계속 감시(참조 보유)해야 하는가?
    문서가 있으면 보유. 문서가 없을 때:
      - Excel: 빈 시작화면(Visible=True)은 계속 감시, 종료 중 좀비(숨김)는 놓아준다.
      - PowerPoint: Visible 이 항상 True 라 구분 불가 → 문서 없으면 놓아준다."""
    cnt = _open_count(app, key)
    if cnt > 0:
        return True
    if cnt < 0:
        return True   # 상태를 못 읽음(일시적) → 일단 유지
    if key == "Excel":
        return _is_visible(app)
    return False


def worker(stop_event):
    pythoncom.CoInitialize()
    log("감시 스레드 시작")
    state = {k: {"app": None, "sink": None, "next": 0.0, "warned": False}
             for (k, _p, _h) in _APPS}

    while not stop_event.is_set():
        try:
            pythoncom.PumpWaitingMessages()
            process_pending(state["Excel"])

            now = time.time()
            for key, progid, handler in _APPS:
                st = state[key]

                if st["sink"] is not None:
                    # 앱이 죽었거나 더 이상 감시할 필요가 없으면(문서 없음/종료 중)
                    # 참조를 놓아준다 → 종료 중인 프로세스가 완전히 종료될 수 있다.
                    if not _is_alive(st["app"]):
                        log("%s: 종료 감지 - 연결 해제" % key)
                        _release_app(st)
                        st["next"] = now + 0.5
                        st["warned"] = False
                        continue
                    if not _should_hold(st["app"], key):
                        log("%s: 문서 없음/종료 중 → 참조 해제(프로세스 종료 허용)" % key)
                        _release_app(st)
                        st["next"] = now + 0.5
                        st["warned"] = False
                        continue

                if st["sink"] is None and now >= st["next"]:
                    app = None
                    stale = False
                    try:
                        app = win32com.client.GetActiveObject(progid)
                    except Exception:
                        app = None
                    if app is not None and not (_is_alive(app)
                                                and _should_hold(app, key)):
                        # 종료 중이거나 문서 없는 낡은 인스턴스 → 붙잡지 않는다
                        app = None
                        stale = True
                    if app is None:
                        # 실제 문서를 가진 인스턴스를 ROT 에서 직접 찾는다
                        app = _find_app_via_rot(key)
                        if app is not None and stale:
                            log("%s: 낡은 인스턴스 무시하고 실제 인스턴스 탐색됨" % key)
                    if app is not None:
                        try:
                            st["sink"] = win32com.client.WithEvents(app, handler)
                            st["app"] = app
                            st["warned"] = False
                            log("%s: 실행 인스턴스에 연결됨" % key)
                        except Exception as e:
                            st["next"] = now + _RETRY_DELAY
                            if not st["warned"]:
                                log("%s: 연결 실패(%.0f초 후 재시도) - %s"
                                    % (key, _RETRY_DELAY, e))
                                st["warned"] = True
                    else:
                        app = None
                        st["next"] = now + 0.5
                        # 낡은 인스턴스만 계속 잡히는 상황은 조용히 넘기지 말고
                        # 주기적으로 남겨 진단할 수 있게 한다.
                        if stale and now - st.get("idle_log", 0) > 60:
                            st["idle_log"] = now
                            log("%s: 감시 대기 중(문서 열린 인스턴스를 찾지 못함)" % key)
        except Exception:
            log("worker 루프 예외\n" + traceback.format_exc())

        stop_event.wait(0.25)

    # 도구 종료 시: 사용자의 Office 는 닫지 않고 이벤트 연결만 해제
    for key in state:
        _release_app(state[key])
    try:
        pythoncom.CoUninitialize()
    except Exception:
        pass
    log("감시 스레드 종료")


# ---------------------------------------------------------------------------
# 트레이 아이콘
# ---------------------------------------------------------------------------
def _icon_label():
    return "FIT" if _fit_mode else str(int(_target_zoom))


def make_app_icon(size=64):
    """앱 아이콘(문서 + 체크). 설정/정보 창의 제목 표시줄에 쓴다.
    tools/make_icon.py 가 만드는 icon.ico 와 같은 모양을 런타임에 그린다."""
    from PIL import Image, ImageDraw

    S = 256
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, S - 1, S - 1], radius=int(S * 0.22),
                        fill=(35, 175, 95, 255))
    left, right = int(S * 0.26), int(S * 0.74)
    top, bottom = int(S * 0.20), int(S * 0.80)
    fold = int(S * 0.16)
    d.polygon([(left, top), (right - fold, top), (right, top + fold),
               (right, bottom), (left, bottom)], fill=(255, 255, 255, 255))
    d.polygon([(right - fold, top), (right, top + fold),
               (right - fold, top + fold)], fill=(214, 232, 221, 255))
    cw = max(2, int(S * 0.055))
    pts = [(int(S * 0.355), int(S * 0.520)),
           (int(S * 0.455), int(S * 0.620)),
           (int(S * 0.650), int(S * 0.395))]
    d.line([pts[0], pts[1]], fill=(33, 160, 88, 255), width=cw)
    d.line([pts[1], pts[2]], fill=(33, 160, 88, 255), width=cw)
    for x, y in pts:
        d.ellipse([x - cw // 2, y - cw // 2, x + cw // 2, y + cw // 2],
                  fill=(33, 160, 88, 255))
    return img.resize((size, size), Image.LANCZOS)


def _set_window_icon(root):
    """Tk 창의 아이콘을 앱 아이콘으로 바꾼다(기본 깃털 아이콘 대체).
    실패해도 창은 정상 동작해야 하므로 조용히 넘어간다."""
    try:
        from PIL import ImageTk
        img = ImageTk.PhotoImage(make_app_icon(64))
        root.iconphoto(True, img)
        root._app_icon_ref = img      # GC 방지용 참조 유지
    except Exception:
        pass


def make_icon_image(label, paused):
    from PIL import Image, ImageDraw, ImageFont

    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    bg = (240, 190, 40, 255) if paused else (33, 150, 83, 255)  # 노랑 / 초록
    d.ellipse([2, 2, size - 2, size - 2], fill=bg)

    fsize = 30 if len(label) <= 2 else 23
    font = None
    for path in (r"C:\Windows\Fonts\arialbd.ttf", r"C:\Windows\Fonts\arial.ttf"):
        try:
            font = ImageFont.truetype(path, fsize)
            break
        except Exception:
            pass
    if font is None:
        font = ImageFont.load_default()

    fill = (60, 50, 0, 255) if paused else (255, 255, 255, 255)
    try:
        bbox = d.textbbox((0, 0), label, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (size - tw) / 2 - bbox[0]
        y = (size - th) / 2 - bbox[1]
    except Exception:
        x, y = 14, 18
    d.text((x, y), label, fill=fill, font=font)
    return img


def _title():
    if _paused:
        state = "일시정지"
    elif _fit_mode:
        state = "창에 맞춤"
    else:
        state = "배율 %d%%" % _target_zoom
    return "ConsultingReady - %s" % state


def _status_text(item):
    if _paused:
        head = "일시정지 중"
    else:
        head = "감시 중"
    mode = "창에 맞춤" if _fit_mode else ("고정 배율 %d%%" % _target_zoom)
    return "%s · %s" % (head, mode)


def _refresh_icon():
    if _icon is not None:
        try:
            _icon.icon = make_icon_image(_icon_label(), _paused)
            _icon.title = _title()
            _icon.update_menu()
        except Exception:
            log("아이콘 갱신 실패\n" + traceback.format_exc())


def apply_settings(zoom=None, fit=None, toast=None,
                   filter_enabled=None, filter_mode=None, filter_keywords=None,
                   stamp_check=None, hide_notes=None,
                   backup_enabled=None, backup_dir=None, backup_keep=None):
    global _target_zoom, _fit_mode, _toast_enabled
    global _filter_enabled, _filter_mode, _filter_keywords, _stamp_check
    global _hide_notes, _backup_enabled, _backup_dir, _backup_keep
    if stamp_check is not None:
        _stamp_check = bool(stamp_check)
    if hide_notes is not None:
        _hide_notes = bool(hide_notes)
    if backup_enabled is not None:
        _backup_enabled = bool(backup_enabled)
    if backup_dir is not None:
        _backup_dir = str(backup_dir).strip()
    if backup_keep is not None:
        try:
            _backup_keep = max(1, min(200, int(backup_keep)))
        except Exception:
            pass
    if zoom is not None:
        _target_zoom = _clamp_zoom(zoom)
    if fit is not None:
        _fit_mode = bool(fit)
    if toast is not None:
        _toast_enabled = bool(toast)
    if filter_enabled is not None:
        _filter_enabled = bool(filter_enabled)
    if filter_mode in ("exclude", "include"):
        _filter_mode = filter_mode
    if filter_keywords is not None:
        _filter_keywords = str(filter_keywords)
    refresh_filter()
    save_config()
    log("설정 변경: 배율 %d%% / 창에맞춤 %s / 완료알림 %s / 필터 %s / 딱지검사 %s"
        " / 메모창접기 %s"
        % (_target_zoom,
           "켜짐" if _fit_mode else "꺼짐",
           "켜짐" if _toast_enabled else "꺼짐",
           _filter_desc(),
           "켜짐" if _stamp_check else "꺼짐",
           "켜짐" if _hide_notes else "꺼짐"))
    log("   백업: %s" % (("켜짐 → %s (원본당 %d개 보관)"
                          % (backup_root(), _backup_keep))
                         if _backup_enabled else "꺼짐"))
    _refresh_icon()


def toggle_pause(icon, item):
    global _paused
    _paused = not _paused
    log("일시정지" if _paused else "재개")
    _refresh_icon()


def toggle_fit(icon, item):
    global _fit_mode
    _fit_mode = not _fit_mode
    save_config()
    log("창에 맞춤 %s" % ("켜짐(배율 고정 비활성화)" if _fit_mode else "꺼짐"))
    _refresh_icon()


def open_settings(icon, item):
    global _settings_open
    if _settings_open:
        return
    _settings_open = True
    threading.Thread(target=_settings_thread, daemon=True).start()


def _settings_thread():
    global _settings_open
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.title("ConsultingReady 설정")
        _set_window_icon(root)
        root.resizable(False, False)
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        PAD = {"padx": 16}
        zoom_var = tk.StringVar(value=str(int(_target_zoom)))
        fit_var = tk.BooleanVar(value=bool(_fit_mode))
        toast_var = tk.BooleanVar(value=bool(_toast_enabled))

        tk.Label(root, text="고정할 배율  (%d ~ %d %%)" % (ZOOM_MIN, ZOOM_MAX),
                 anchor="w", font=("Malgun Gothic", 9, "bold")
                 ).pack(fill="x", pady=(16, 4), **PAD)
        entry = tk.Entry(root, textvariable=zoom_var, width=8, justify="right",
                         font=("Malgun Gothic", 11))
        entry.pack(anchor="w", **PAD)

        def sync_entry(*_a):
            entry.configure(state="disabled" if fit_var.get() else "normal")

        tk.Checkbutton(root, variable=fit_var, command=sync_entry, anchor="w",
                       justify="left", wraplength=320,
                       text="창에 맞춤으로 저장 (배율 고정 대신)"
                       ).pack(fill="x", pady=(12, 0), **PAD)
        tk.Checkbutton(root, variable=toast_var, anchor="w",
                       text="저장 완료 알림(토스트) 표시"
                       ).pack(fill="x", **PAD)

        stamp_var = tk.BooleanVar(value=bool(_stamp_check))
        tk.Checkbutton(root, variable=stamp_var, anchor="w", justify="left",
                       wraplength=330,
                       text="PowerPoint 딱지 검사 (기울어진 도형이 남아있으면 알림)"
                       ).pack(fill="x", **PAD)

        notes_var = tk.BooleanVar(value=bool(_hide_notes))
        tk.Checkbutton(root, variable=notes_var, anchor="w", justify="left",
                       wraplength=330,
                       text="PowerPoint 메모 창 접고 저장 (발표 준비 상태)"
                       ).pack(fill="x", **PAD)

        tk.Label(root, fg="#666666", anchor="w", justify="left", wraplength=360,
                 text="※ 저장 후 닫을 때 각 시트의 커서를 A1 으로 이동합니다."
                 ).pack(fill="x", pady=(10, 0), **PAD)

        # ---------------- 수정 전 원본 백업 ----------------
        tk.Frame(root, height=1, bg="#d9d9d9").pack(fill="x", pady=12, **PAD)

        b_en = tk.BooleanVar(value=bool(_backup_enabled))
        b_dir = tk.StringVar(value=(_backup_dir.strip()
                                    or os.path.join(_DATA_DIR, "backup")))
        b_keep = tk.StringVar(value=str(int(_backup_keep)))

        tk.Checkbutton(root, variable=b_en, anchor="w",
                       text="수정 전 원본을 백업",
                       font=("Malgun Gothic", 9, "bold"),
                       command=lambda: sync_backup()).pack(fill="x", **PAD)

        brow = tk.Frame(root)
        brow.pack(fill="x", **PAD)
        b_entry = tk.Entry(brow, textvariable=b_dir, font=("Malgun Gothic", 9))
        b_entry.pack(side="left", fill="x", expand=True)

        def pick_dir():
            try:
                from tkinter import filedialog
                d = filedialog.askdirectory(parent=root, title="백업 폴더 선택",
                                            mustexist=True)
                if d:
                    b_dir.set(os.path.normpath(d))
            except Exception:
                pass
        b_btn = tk.Button(brow, text="찾아보기...", command=pick_dir)
        b_btn.pack(side="left", padx=(6, 0))

        krow = tk.Frame(root)
        krow.pack(fill="x", **PAD)
        tk.Label(krow, text="원본 1개당 보관 개수:", anchor="w",
                 font=("Malgun Gothic", 9)).pack(side="left")
        b_keep_entry = tk.Entry(krow, textvariable=b_keep, width=5,
                                justify="right", font=("Malgun Gothic", 9))
        b_keep_entry.pack(side="left", padx=(6, 0))

        b_hint = tk.Label(root, anchor="w", justify="left", wraplength=360,
                          fg="#666666", font=("Malgun Gothic", 8),
                          text="※ 폴더가 없으면 자동으로 만듭니다.\n"
                               "※ 백업에 실패하면 파일을 수정하지 않고 그대로 닫습니다.\n"
                               "※ SharePoint/OneDrive 온라인 문서는 서버 버전 기록이 "
                               "있어 백업하지 않습니다.")
        b_hint.pack(fill="x", **PAD)

        def sync_backup(*_a):
            st = "normal" if b_en.get() else "disabled"
            for w in (b_entry, b_btn, b_keep_entry):
                w.configure(state=st)

        # ---------------- 파일명 키워드 필터 ----------------
        tk.Frame(root, height=1, bg="#d9d9d9").pack(fill="x", pady=12, **PAD)

        f_en = tk.BooleanVar(value=bool(_filter_enabled))
        f_mode = tk.StringVar(value=_filter_mode)
        f_pat = tk.StringVar(value=_filter_keywords)

        tk.Checkbutton(root, variable=f_en, anchor="w",
                       text="파일명 키워드 필터 사용",
                       font=("Malgun Gothic", 9, "bold"),
                       command=lambda: sync_filter()).pack(fill="x", **PAD)

        rb1 = tk.Radiobutton(root, variable=f_mode, value="exclude", anchor="w",
                             text="제외 — 키워드가 들어간 파일은 건드리지 않음")
        rb2 = tk.Radiobutton(root, variable=f_mode, value="include", anchor="w",
                             text="허용 — 키워드가 들어간 파일만 처리")
        rb1.pack(fill="x", padx=30)
        rb2.pack(fill="x", padx=30)

        pat_entry = tk.Entry(root, textvariable=f_pat, font=("Malgun Gothic", 10))
        pat_entry.pack(fill="x", pady=(6, 2), **PAD)
        hint = tk.Label(root, anchor="w", justify="left", wraplength=360,
                        fg="#666666", font=("Malgun Gothic", 8),
                        text="예) 구글, 제미나이   (쉼표로 구분 · 하나라도 들어 있으면 일치 · "
                             "띄어쓰기와 대소문자는 무시)")
        hint.pack(fill="x", **PAD)
        status = tk.Label(root, anchor="w", justify="left", wraplength=360,
                          fg="#666666", font=("Malgun Gothic", 8), text="")
        status.pack(fill="x", **PAD)

        def check_pattern(*_a):
            if not f_en.get():
                status.configure(text="", fg="#666666")
                return
            words = [w.strip() for w in _KW_SPLIT.split(f_pat.get()) if w.strip()]
            if not words:
                status.configure(text="키워드가 비어 있어 필터가 적용되지 않습니다.",
                                 fg="#b06000")
                return
            status.configure(text="키워드 %d개: %s" % (len(words), " / ".join(words)),
                             fg="#1a7f37")

        def sync_filter(*_a):
            on = f_en.get()
            st = "normal" if on else "disabled"
            for w in (rb1, rb2, pat_entry):
                w.configure(state=st)
            check_pattern()

        f_pat.trace_add("write", check_pattern)

        def on_save():
            # 백업 설정 검증
            bd = b_dir.get().strip()
            try:
                bk = int(b_keep.get().strip())
            except Exception:
                bk = -1
            if b_en.get():
                if not bd:
                    messagebox.showwarning("입력 확인",
                                           "백업 폴더를 지정해 주세요.", parent=root)
                    return
                if not os.path.isdir(bd):
                    # 폴더가 없으면 만들어 준다. 만들 수 없는 경로일 때만 막는다.
                    try:
                        os.makedirs(bd, exist_ok=True)
                        log("백업 폴더 생성: %s" % bd)
                    except Exception as e:
                        messagebox.showwarning(
                            "입력 확인",
                            "백업 폴더를 만들 수 없습니다.\n경로를 확인해 주세요.\n\n%s\n\n%s"
                            % (bd, e), parent=root)
                        return
                if not (1 <= bk <= 200):
                    messagebox.showwarning("입력 확인",
                                           "보관 개수는 1 ~ 200 사이여야 합니다.",
                                           parent=root)
                    return
            fkw = dict(filter_enabled=f_en.get(),
                       filter_mode=f_mode.get(),
                       filter_keywords=f_pat.get().strip(),
                       stamp_check=stamp_var.get(),
                       hide_notes=notes_var.get(),
                       backup_enabled=b_en.get(),
                       backup_dir=bd,
                       backup_keep=bk if bk > 0 else 20)
            if fit_var.get():
                apply_settings(fit=True, toast=toast_var.get(), **fkw)
            else:
                raw = zoom_var.get().strip().rstrip("%").strip()
                try:
                    val = int(round(float(raw)))
                except Exception:
                    messagebox.showwarning("입력 확인",
                                           "배율은 숫자로 입력하세요.", parent=root)
                    return
                if not (ZOOM_MIN <= val <= ZOOM_MAX):
                    messagebox.showwarning(
                        "입력 확인",
                        "배율은 %d ~ %d 사이여야 합니다." % (ZOOM_MIN, ZOOM_MAX),
                        parent=root)
                    return
                apply_settings(zoom=val, fit=False, toast=toast_var.get(), **fkw)
            root.destroy()

        btns = tk.Frame(root)
        btns.pack(fill="x", pady=16, **PAD)
        # 오른쪽부터 pack 되므로, 화면에는 [저장] [취소] 순으로 놓인다.
        tk.Button(btns, text="취소", width=9,
                  command=root.destroy).pack(side="right")
        tk.Button(btns, text="저장", width=9,
                  command=on_save).pack(side="right", padx=(0, 8))

        sync_entry()
        sync_filter()
        sync_backup()
        root.update_idletasks()
        w, h = root.winfo_width(), root.winfo_height()
        x = (root.winfo_screenwidth() - w) // 2
        y = (root.winfo_screenheight() - h) // 2
        root.geometry("+%d+%d" % (x, y))

        entry.focus_set()
        entry.selection_range(0, "end")
        root.bind("<Return>", lambda e: on_save())
        root.bind("<Escape>", lambda e: root.destroy())
        root.mainloop()
    except Exception:
        log("설정창 오류\n" + traceback.format_exc())
    finally:
        _settings_open = False


def open_about(icon, item):
    global _about_open
    if _about_open:
        return
    _about_open = True
    threading.Thread(target=_about_thread, daemon=True).start()


def _about_thread():
    global _about_open
    try:
        import tkinter as tk

        root = tk.Tk()
        root.title("프로그램 정보")
        _set_window_icon(root)
        root.resizable(False, False)
        try:
            root.attributes("-topmost", True)
        except Exception:
            pass

        PAD = {"padx": 20}
        tk.Label(root, text=APP_NAME, anchor="w",
                 font=("Malgun Gothic", 13, "bold")
                 ).pack(fill="x", pady=(18, 0), **PAD)
        tk.Label(root, text="버전 %s" % APP_VERSION, anchor="w", fg="#666666",
                 font=("Malgun Gothic", 9)).pack(fill="x", **PAD)
        tk.Label(root, text=APP_DESC, anchor="w", justify="left", fg="#333333",
                 font=("Malgun Gothic", 9)).pack(fill="x", pady=(12, 0), **PAD)

        tk.Frame(root, height=1, bg="#d9d9d9").pack(fill="x", pady=14, **PAD)

        tk.Label(root, text="제작자", anchor="w",
                 font=("Malgun Gothic", 9, "bold")).pack(fill="x", **PAD)
        tk.Label(root, text=APP_AUTHOR, anchor="w", font=("Malgun Gothic", 9)
                 ).pack(fill="x", **PAD)
        tk.Label(root, text=APP_EMAIL, anchor="w", fg="#1a5fb4",
                 font=("Malgun Gothic", 9)).pack(fill="x", **PAD)

        def copy_mail():
            try:
                root.clipboard_clear()
                root.clipboard_append(APP_EMAIL)
                btn_copy.configure(text="복사됨")
                root.after(1200, lambda: btn_copy.configure(text="메일 주소 복사"))
            except Exception:
                pass

        btns = tk.Frame(root)
        btns.pack(fill="x", pady=(16, 18), **PAD)
        tk.Button(btns, text="확인", width=9,
                  command=root.destroy).pack(side="right")
        btn_copy = tk.Button(btns, text="메일 주소 복사", width=14,
                             command=copy_mail)
        btn_copy.pack(side="right", padx=(0, 8))

        root.update_idletasks()
        w, h = root.winfo_width(), root.winfo_height()
        root.geometry("+%d+%d" % ((root.winfo_screenwidth() - w) // 2,
                                  (root.winfo_screenheight() - h) // 2))
        root.bind("<Escape>", lambda e: root.destroy())
        root.mainloop()
    except Exception:
        log("정보창 오류\n" + traceback.format_exc())
    finally:
        _about_open = False


def main():
    global _icon, _target_zoom, _fit_mode, _toast_enabled, _mutex
    global _filter_enabled, _filter_mode, _filter_keywords, _stamp_check
    global _hide_notes, _backup_enabled, _backup_dir, _backup_keep

    # 단일 실행 보장: mutex 핸들을 전역에 보관해 프로세스 수명 동안 유지한다.
    _mutex = win32event.CreateMutex(None, False, "Global\\ConsultingReadySingleton")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        log("이미 실행 중 - 종료")
        return

    (_target_zoom, _fit_mode, _toast_enabled,
     _filter_enabled, _filter_mode, _filter_keywords,
     _stamp_check, _hide_notes,
     _backup_enabled, _backup_dir, _backup_keep) = load_config()
    refresh_filter()
    log("=== ConsultingReady 시작 (%s, 완료알림 %s, 필터 %s, 딱지검사 %s,"
        " 메모창접기 %s) ==="
        % ("창에 맞춤" if _fit_mode else "배율 %d%%" % _target_zoom,
           "켜짐" if _toast_enabled else "꺼짐",
           _filter_desc(),
           "켜짐" if _stamp_check else "꺼짐",
           "켜짐" if _hide_notes else "꺼짐"))
    log("   백업: %s" % (("켜짐 → %s (원본당 %d개 보관)"
                          % (backup_root(), _backup_keep))
                         if _backup_enabled else "꺼짐"))

    stop_event = threading.Event()
    t = threading.Thread(target=worker, args=(stop_event,), daemon=True)
    t.start()
    toast_t = threading.Thread(target=_toast_loop, args=(stop_event,), daemon=True)
    toast_t.start()

    try:
        import pystray

        def on_exit(icon, item):
            log("트레이 메뉴에서 종료 요청")
            stop_event.set()
            icon.stop()

        def on_open_log(icon, item):
            try:
                os.startfile(_LOG_PATH)
            except Exception:
                pass

        menu = pystray.Menu(
            pystray.MenuItem(_status_text, None, enabled=False),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("설정 (배율 · 알림)...", open_settings),
            pystray.MenuItem(
                "일시정지", toggle_pause, checked=lambda item: _paused
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("로그 열기", on_open_log),
            pystray.MenuItem("프로그램 정보", open_about),
            pystray.MenuItem("종료", on_exit),
        )
        _icon = pystray.Icon(
            "ConsultingReady",
            make_icon_image(_icon_label(), _paused),
            _title(),
            menu,
        )
        _icon.run()
    except Exception:
        log("트레이 초기화 실패, 트레이 없이 실행\n" + traceback.format_exc())
        try:
            while not stop_event.is_set():
                time.sleep(1)
        except KeyboardInterrupt:
            pass

    stop_event.set()
    t.join(timeout=5)
    log("=== ConsultingReady 종료 ===")


if __name__ == "__main__":
    main()
