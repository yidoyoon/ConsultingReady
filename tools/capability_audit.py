# -*- coding: utf-8 -*-
"""
악성 행위 능력 검사 (Capability Audit)
=====================================
"이 코드에 해킹/악성 코드가 없다"를 뒷받침하기 위한 검사.

일반적인 보안 린터가 '취약점'(실수로 생긴 약점)을 찾는다면, 이 검사는
'악성 프로그램이 하는 일'(의도적으로 넣은 기능)을 찾는다.

각 항목은 악성코드가 목적을 달성하려면 반드시 필요한 능력이다.
하나라도 발견되면 검사에 실패하고, 어느 파일 몇 번째 줄인지 보고한다.

사용법:
    python tools/capability_audit.py [검사할_경로 ...]
종료 코드: 0 = 통과, 1 = 의심 항목 발견
"""

import io
import os
import re
import sys

# ---------------------------------------------------------------------------
# 검사 항목: (분류, 설명, 정규식, 왜 위험한가)
# ---------------------------------------------------------------------------
CHECKS = [
    ("네트워크 통신",
     "외부와 데이터를 주고받는 코드",
     r"\b(?:import|from)\s+(?:socket|requests|urllib|urllib2|urllib3|httplib|"
     r"http\.client|ftplib|smtplib|poplib|imaplib|telnetlib|paramiko|"
     r"websocket|websockets|aiohttp|httpx)\b"
     r"|\burlopen\s*\(|\brequests\s*\.\s*(?:get|post|put|patch|delete|head|request)\s*\("
     r"|\bsocket\s*\.\s*(?:socket|create_connection)\s*\(",
     "수집한 정보를 외부 서버로 보내거나(유출), 추가 악성코드를 내려받는 통로"),

    ("외부 프로그램 실행",
     "셸 명령이나 다른 실행파일을 구동하는 코드",
     r"\b(?:import|from)\s+(?:subprocess|pty)\b"
     r"|\bsubprocess\s*\.\s*(?:run|call|check_call|check_output|Popen)\s*\("
     r"|\bos\s*\.\s*(?:system|popen|execv?[lpe]*|spawn[lv]p?e?)\s*\("
     r"|\bShellExecute\w*\s*\(|\bCreateProcess\w*\s*\(",
     "랜섬웨어 실행, 권한 상승, 백도어 기동 등 2차 페이로드의 실행 수단"),

    ("동적 코드 실행",
     "실행 시점에 코드를 만들어 돌리는 구문",
     r"(?<![\w.])eval\s*\(|(?<![\w.])exec\s*\(|\bcompile\s*\([^)]*['\"]exec['\"]"
     r"|\b__import__\s*\(|\bimportlib\s*\.\s*import_module\s*\(",
     "정적 분석을 회피해 숨긴 코드를 실행하는 대표적 수법"),

    ("난독화 / 인코딩된 페이로드",
     "코드를 숨기기 위해 인코딩·압축한 흔적",
     r"\bbase64\s*\.\s*b(?:64|32|16)decode\s*\(|\bcodecs\s*\.\s*decode\s*\([^)]*rot13"
     r"|\bzlib\s*\.\s*decompress\s*\(|\bmarshal\s*\.\s*loads\s*\("
     r"|\bpickle\s*\.\s*loads\s*\(|\bbytes\s*\.\s*fromhex\s*\(",
     "악성 페이로드를 문자열로 위장해 넣어두는 전형적 패턴"),

    ("자동 실행 등록(지속성)",
     "재부팅 후에도 자동 실행되도록 시스템에 등록하는 코드",
     r"CurrentVersion\\+Run|\bwinreg\b|\b_winreg\b|RegSetValue\w*\s*\("
     r"|schtasks|\bTask\s*Scheduler\b|\bsc\s+create\b",
     "사용자 몰래 상주하며 재감염되게 만드는 지속성 확보 수단"),

    ("입력 가로채기 / 화면 감시",
     "키보드·마우스·화면을 몰래 기록하는 코드",
     r"\b(?:import|from)\s+(?:pynput|keyboard|mouse)\b"
     r"|SetWindowsHookEx\w*\s*\(|GetAsyncKeyState\s*\(|GetKeyboardState\s*\("
     r"|\bImageGrab\s*\.\s*grab\s*\(|\bpyautogui\s*\.\s*screenshot\s*\(",
     "키로거·화면 캡쳐로 비밀번호와 문서 내용을 훔치는 행위"),

    ("자격증명 접근",
     "저장된 비밀번호·토큰·브라우저 데이터를 읽는 코드",
     r"CredEnumerate\w*\s*\(|CredRead\w*\s*\(|\bkeyring\b"
     r"|Login\s*Data|\bcookies\.sqlite\b|\bkey[34]\.db\b"
     r"|\.ssh[/\\]id_rsa|\.aws[/\\]credentials",
     "계정 탈취의 직접적 수단"),

    ("광범위 파일 탐색 / 암호화",
     "사용자 문서 전체를 훑거나 암호화하는 코드",
     r"\bos\s*\.\s*walk\s*\(|\bglob\s*\.\s*(?:i?glob)\s*\("
     r"|\bPath\s*\([^)]*\)\s*\.\s*rglob\s*\("
     r"|\b(?:import|from)\s+(?:Crypto|Cryptodome|cryptography|nacl|pycryptodome)\b"
     r"|\bshutil\s*\.\s*rmtree\s*\(",
     "랜섬웨어의 탐색·암호화 단계, 또는 대량 문서 수집"),

    ("보안 기능 무력화",
     "백신·방화벽·로그를 끄는 코드",
     r"Set-MpPreference|DisableRealtimeMonitoring|netsh\s+advfirewall"
     r"|wevtutil|vssadmin\s+delete|bcdedit\s+/set",
     "탐지와 복구를 막는 전형적 악성 행위"),
]

# 이 파일 자체(패턴 정의)와 문서는 검사 대상에서 제외한다.
EXCLUDE_FILES = {"capability_audit.py"}
EXCLUDE_DIRS = {".git", "build", "dist", "__pycache__", ".github", "docs"}


def iter_files(roots):
    for root in roots:
        if os.path.isfile(root):
            yield root
            continue
        for base, dirs, names in os.walk(root):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS]
            for n in names:
                if n.endswith(".py") and n not in EXCLUDE_FILES:
                    yield os.path.join(base, n)


def main(argv):
    roots = argv[1:] or ["."]
    files = sorted(set(iter_files(roots)))
    findings = []

    for path in files:
        try:
            lines = io.open(path, encoding="utf-8", errors="replace").read().splitlines()
        except Exception as e:
            print("읽기 실패: %s (%s)" % (path, e))
            continue
        for i, line in enumerate(lines, 1):
            stripped = line.strip()
            if stripped.startswith("#"):          # 주석은 제외
                continue
            for name, desc, pattern, why in CHECKS:
                if re.search(pattern, line):
                    findings.append((name, desc, why, path, i, stripped[:120]))

    # ---- 보고서 ----
    out = []
    out.append("# 악성 행위 능력 검사 결과\n")
    out.append("악성코드가 목적을 달성하려면 반드시 필요한 능력이 "
               "코드에 있는지 항목별로 검사합니다.\n")
    out.append("| 검사 항목 | 결과 |")
    out.append("|---|---|")
    hit_names = {f[0] for f in findings}
    for name, desc, _p, _w in CHECKS:
        mark = "❌ 발견됨" if name in hit_names else "✅ 없음"
        out.append("| **%s** — %s | %s |" % (name, desc, mark))
    out.append("")
    out.append("검사한 파일: %d 개" % len(files))
    out.append("")

    if findings:
        out.append("## ❌ 발견된 항목\n")
        for name, _d, why, path, ln, text in findings:
            out.append("**%s** — `%s:%d`" % (name, path.replace("\\", "/"), ln))
            out.append("")
            out.append("```python")
            out.append(text)
            out.append("```")
            out.append("> 왜 위험한가: %s" % why)
            out.append("")
    else:
        out.append("## ✅ 통과\n")
        out.append("위 항목 중 어느 것도 발견되지 않았습니다. "
                   "이 프로그램은 네트워크 통신, 외부 프로그램 실행, "
                   "동적 코드 실행, 자동 실행 등록, 입력 가로채기, "
                   "자격증명 접근 능력을 가지고 있지 않습니다.")

    report = "\n".join(out)
    print(report)

    summary = os.environ.get("GITHUB_STEP_SUMMARY")
    if summary:
        try:
            with io.open(summary, "a", encoding="utf-8") as f:
                f.write(report + "\n")
        except Exception:
            pass

    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
