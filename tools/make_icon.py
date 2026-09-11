# -*- coding: utf-8 -*-
"""
앱 아이콘 생성기
================
ConsultingReady 의 아이콘을 코드로 그려 docs/icon.ico / docs/icon.png 를 만든다.

출처를 알 수 없는 바이너리를 저장소에 넣지 않기 위해, 아이콘도 소스에서
재생성할 수 있게 했다. (보안 검증 취지와 동일한 이유)

디자인 의도
  - 초록 둥근 사각형  : 트레이 아이콘과 같은 색 계열 (감시 중 = 정상)
  - 흰 문서 + 접힌 모서리 : Excel / PowerPoint 문서
  - 초록 체크        : "정리·점검이 끝나 내보낼 준비가 됐다"

사용법:  python tools/make_icon.py
"""

import os
import sys

from PIL import Image, ImageDraw

S = 1024                      # 큰 크기로 그린 뒤 축소해 가장자리를 매끄럽게
GREEN_TOP = (46, 204, 113)
GREEN_BOTTOM = (24, 146, 80)
WHITE = (255, 255, 255)
CHECK = (33, 160, 88)

OUT_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "docs")
SIZES = [16, 24, 32, 48, 64, 128, 256]


def rounded_rect_mask(size, radius):
    m = Image.new("L", (size, size), 0)
    d = ImageDraw.Draw(m)
    d.rounded_rectangle([0, 0, size - 1, size - 1], radius=radius, fill=255)
    return m


def vertical_gradient(size, top, bottom):
    g = Image.new("RGB", (1, size))
    for y in range(size):
        t = y / float(size - 1)
        g.putpixel((0, y), tuple(int(top[i] + (bottom[i] - top[i]) * t)
                                 for i in range(3)))
    return g.resize((size, size))


def draw_icon():
    # 배경: 세로 그라데이션 + 둥근 모서리
    bg = vertical_gradient(S, GREEN_TOP, GREEN_BOTTOM).convert("RGBA")
    bg.putalpha(rounded_rect_mask(S, int(S * 0.22)))

    layer = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    d = ImageDraw.Draw(layer)

    # 문서 본체 (모서리 하나가 접힌 형태)
    left, right = int(S * 0.26), int(S * 0.74)
    top, bottom = int(S * 0.20), int(S * 0.80)
    fold = int(S * 0.16)                       # 접힌 모서리 크기
    body = [
        (left, top),
        (right - fold, top),
        (right, top + fold),
        (right, bottom),
        (left, bottom),
    ]
    d.polygon(body, fill=WHITE)

    # 접힌 모서리 (살짝 어두운 흰색으로 입체감)
    d.polygon([(right - fold, top), (right, top + fold), (right - fold, top + fold)],
              fill=(214, 232, 221, 255))

    # 문서 안의 체크 표시
    cw = int(S * 0.055)                        # 선 굵기
    p1 = (int(S * 0.355), int(S * 0.520))
    p2 = (int(S * 0.455), int(S * 0.620))
    p3 = (int(S * 0.650), int(S * 0.395))
    d.line([p1, p2], fill=CHECK, width=cw)
    d.line([p2, p3], fill=CHECK, width=cw)
    for p in (p1, p2, p3):                     # 선 끝을 둥글게
        d.ellipse([p[0] - cw // 2, p[1] - cw // 2,
                   p[0] + cw // 2, p[1] + cw // 2], fill=CHECK)

    # 문서 위쪽 텍스트 줄 두 개 (문서임을 더 분명히)
    line_h = int(S * 0.028)
    for i, (x0, x1) in enumerate(((0.345, 0.600), (0.345, 0.540))):
        y = int(S * (0.300 + i * 0.070))
        d.rounded_rectangle([int(S * x0), y, int(S * x1), y + line_h],
                            radius=line_h // 2, fill=(208, 219, 213, 255))

    return Image.alpha_composite(bg, layer)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    icon = draw_icon()

    png_path = os.path.join(OUT_DIR, "icon.png")
    icon.resize((256, 256), Image.LANCZOS).save(png_path)

    ico_path = os.path.join(OUT_DIR, "icon.ico")
    frames = [icon.resize((s, s), Image.LANCZOS) for s in SIZES]
    frames[-1].save(ico_path, format="ICO",
                    sizes=[(s, s) for s in SIZES])

    print("생성 완료")
    print("  %s" % png_path)
    print("  %s  (크기: %s)" % (ico_path, ", ".join("%dx%d" % (s, s) for s in SIZES)))
    return 0


if __name__ == "__main__":
    sys.exit(main())
