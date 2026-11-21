#!/usr/bin/env python3
"""
노선이미지 생성기 (macOS / Windows / Linux 공용)
EBASE BUS_Controller 스타일 192x160 LED 행선판 BMP 생성기

필요: Python 3.9+ (tkinter 포함), Pillow, openpyxl
  pip3 install pillow openpyxl numpy
실행: python3 bus_sign_maker.py
"""
import json
import os
import platform
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

APP_VERSION = "2.0 (2026-11-19)"
APP_DIR = Path.home() / "BusSignMaker"
OUT_DIR = APP_DIR / "output"
PRESET_FILE = APP_DIR / "presets.json"

# ───────────────────────── 폰트 ─────────────────────────
# 앱에 넣어 둔 폰트를 최우선으로 사용 → 어느 컴퓨터에서든 결과가 픽셀 단위로 같음
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BUNDLED_FONT = BASE_DIR / "fonts" / "NanumGothic-Regular.ttf"
BUNDLED_BOLD = BASE_DIR / "fonts" / "NanumGothic-Bold.ttf"
FONT_CANDIDATES = [
    str(BUNDLED_FONT),
    # macOS
    "/System/Library/Fonts/AppleSDGothicNeo.ttc",
    "/System/Library/Fonts/Supplemental/AppleGothic.ttf",
    "/Library/Fonts/NanumGothic.ttf",
    str(Path.home() / "Library/Fonts/NanumGothic.ttf"),
    # Windows
    "C:/Windows/Fonts/malgun.ttf",
    # Linux
    "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    "/usr/share/fonts/truetype/nanum/NanumGothic.ttf",
]
NUM_FONT_CANDIDATES = [
    str(BUNDLED_BOLD),
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "/System/Library/Fonts/Helvetica.ttc",
    "C:/Windows/Fonts/arialbd.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
]


def find_font(cands):
    for p in cands:
        if os.path.exists(p):
            return p
    return None


TEXT_FONT = find_font(FONT_CANDIDATES)
NUM_FONT = find_font(NUM_FONT_CANDIDATES) or TEXT_FONT
_font_cache = {}


def get_font(path, size):
    key = (path, size)
    if key not in _font_cache:
        # 배치 엔진을 BASIC으로 고정: Raqm이 있는 PC(리눅스)와 없는 PC(맥)에서 글자 간격이 달라지는 것 방지
        _font_cache[key] = (ImageFont.truetype(path, size, layout_engine=ImageFont.Layout.BASIC)
                            if path else ImageFont.load_default())
    return _font_cache[key]


FONT_EXTS = (".ttf", ".ttc", ".otf")
FONT_DIRS = [
    APP_DIR / "fonts",                       # 여기에 폰트 넣으면 목록에 자동 표시
    Path.home() / "Library/Fonts",
    Path("/Library/Fonts"),
    Path("/System/Library/Fonts"),
    Path("/System/Library/Fonts/Supplemental"),
    Path("C:/Windows/Fonts"),
    Path("/usr/share/fonts"),
]


def scan_fonts():
    """설치된 폰트 {표시이름: 경로}. 앱 fonts 폴더가 맨 앞."""
    found = {}
    for d in FONT_DIRS:
        if not d.exists():
            continue
        it = d.rglob("*") if d == Path("/usr/share/fonts") else d.iterdir()
        for p in sorted(it):
            if p.suffix.lower() in FONT_EXTS and p.name not in found:
                found[p.name] = str(p)
    return found


# 현재 렌더에 쓰이는 폰트 설정 (render_route가 스타일에서 채움)
CFG = {"text": TEXT_FONT, "num": NUM_FONT, "pixel": False, "thresh": 110, "stroke": 1, "num_stroke": 3, "led_digits": True}
# 글자 획 두께(픽셀). 0 = 폰트 모양 그대로(획 두께 들쭉날쭉)
WEIGHTS = {"(기본)": 1, "1픽셀": 1, "2픽셀(큰 글자용)": 2, "폰트 그대로": 0}
NUM_WEIGHTS = {"(기본)": 3, "원본형": 3, "균일 2픽셀": 2, "균일 3픽셀": 3, "폰트 그대로": 0}


def text_font():
    return CFG["text"] or TEXT_FONT


def num_font():
    return CFG["num"] or text_font()


# ───────────────────────── 색 / 프리셋 ─────────────────────────
COLORS = {
    "(기본)": None,
    "노랑": (255, 230, 0),
    "흰색": (255, 255, 255),
    "빨강": (255, 40, 40),
    "초록": (40, 255, 60),
    "주황": (255, 140, 0),
    "하늘": (80, 200, 255),
}
STYLE_KEYS = [  # (키, 라벨, 기본값)
    ("num_big", "큰번호색", "노랑"),
    ("start", "기점색", "흰색"),
    ("end", "종점색", "흰색"),
    ("via", "경유지색", "흰색"),
    ("sep", "구분자색", "노랑"),
    ("arrow", "화살표색", "노랑"),
    ("num_small", "작은번호색", "노랑"),   # 각 줄 왼쪽 작은 번호
    ("arrow_shape", "화살표", "◆"),
    ("align", "정렬", "가운데"),
    ("weight", "글자획", "1픽셀"),
    ("num_weight", "번호획", "원본형"),
]
# 폰트 설정도 프리셋에 같이 저장됨 (빈 값 = 자동)
FONT_KEYS = ["text_font", "num_font", "pixel_mode"]
ARROW_SHAPES = ["(기본)", "◆", "→", "▶", "↔"]
ALIGNS = ["(기본)", "가운데", "왼쪽"]


def resolve_style(style):
    out = {}
    for k, _, default in STYLE_KEYS:
        v = (style or {}).get(k, "(기본)")
        if v in (None, "", "(기본)"):
            v = default
        out[k] = v
    for k in FONT_KEYS:
        out[k] = (style or {}).get(k) or ""
    out["kinds"] = (style or {}).get("kinds") or {}
    return out


def col(name):
    return COLORS.get(name) or (255, 255, 255)


# ───────────────────────── 렌더러 ─────────────────────────
W, H = 192, 160
RENDER_SIZE = 64  # 고해상도로 그린 뒤 축소 → 픽셀화


def render_segments(segments, target_h, max_w, font_path=None, stroke=None):
    """[(텍스트, RGB), ...] 한 줄을 target_h 높이·max_w 이하 폭의 RGBA 이미지로."""
    is_num = font_path is not None and font_path == num_font()
    font_path = font_path or text_font()
    segs = [(t, c) for t, c in segments if t]
    if not segs:
        return Image.new("RGBA", (1, target_h), (0, 0, 0, 0))
    if is_num and CFG["led_digits"] and CFG["num_stroke"] > 0 and all(
            ch in LED_CHARS for t, _ in segs for ch in t):
        return render_led_digits(segs, target_h, max_w, CFG["num_stroke"])
    if CFG["pixel"]:
        return render_segments_pixel(segs, target_h, max_w, font_path)
    font = get_font(font_path, RENDER_SIZE)
    total = int(sum(font.getlength(t) for t, _ in segs)) + 8
    canvas_h = RENDER_SIZE * 2
    big = Image.new("RGBA", (total, canvas_h), (0, 0, 0, 0))
    d = ImageDraw.Draw(big)
    x = 0.0
    for t, c in segs:
        d.text((x, RENDER_SIZE * 0.4), t, font=font, fill=c + (255,))
        x += font.getlength(t)
    bbox = big.getbbox()
    if not bbox:
        return Image.new("RGBA", (1, target_h), (0, 0, 0, 0))
    # 세로는 폰트 기준선(글자마다 높이 다르지 않게)으로 자름
    # 기준 글자("힣Ag|")의 실제 잉크 높이로 자름 → 여백 없이 줄 높이를 꽉 채움
    _, rt, _, rb = font.getbbox("힣갈Ag", anchor="la")
    top = int(RENDER_SIZE * 0.4 + rt)
    bot = int(RENDER_SIZE * 0.4 + rb)
    top, bot = min(top, bbox[1]), max(bot, bbox[3])
    big = big.crop((bbox[0], top, bbox[2], bot))
    scale = target_h / big.height
    w = max(1, min(max_w, round(big.width * scale)))
    if stroke is None:
        stroke = CFG["num_stroke"] if is_num else CFG["stroke"]
    if stroke and stroke > 0:
        return uniform_stroke_from_big(big, w, target_h, stroke)
    small = big.resize((w, target_h), Image.LANCZOS)
    # 폰트 그대로: 단순 이진화
    r, g, b, a = small.split()
    th = CFG["thresh"]
    a = a.point(lambda v: 255 if v > th else 0)
    return Image.merge("RGBA", (r, g, b, a))


def _thin(m):
    """Zhang-Suen 세선화 (numpy). m: bool 2D 배열 → 1픽셀 중심선."""
    import numpy as np
    m = np.pad(m.astype(np.uint8), 1)
    while True:
        changed = False
        for step in (0, 1):
            p2 = np.roll(m, 1, 0); p6 = np.roll(m, -1, 0)
            p8 = np.roll(m, 1, 1); p4 = np.roll(m, -1, 1)
            p3 = np.roll(p2, -1, 1); p9 = np.roll(p2, 1, 1)
            p5 = np.roll(p6, -1, 1); p7 = np.roll(p6, 1, 1)
            nb = [p2, p3, p4, p5, p6, p7, p8, p9]
            B = sum(n.astype(np.int16) for n in nb)
            A = sum(((nb[i] == 0) & (nb[(i + 1) % 8] == 1)).astype(np.int16) for i in range(8))
            if step == 0:
                c = (p2 * p4 * p6 == 0) & (p4 * p6 * p8 == 0)
            else:
                c = (p2 * p4 * p8 == 0) & (p2 * p6 * p8 == 0)
            kill = (m == 1) & (B >= 2) & (B <= 6) & (A == 1) & c
            if kill.any():
                m[kill] = 0
                changed = True
        if not changed:
            break
    return m[1:-1, 1:-1].astype(bool)


def _thin_seq(m):
    """작은 이미지용 순차 세선화: 2px 겹친 곳만 1px로, 짧은 획은 지우지 않음."""
    import numpy as np
    h, w = m.shape
    g = np.pad(m.astype(np.uint8), 1).tolist()
    order = ((0, -1), (1, 0), (0, 1), (-1, 0))  # 위·오른쪽·아래·왼쪽 테두리부터 깎기
    changed = True
    while changed:
        changed = False
        for dx, dy in order:
            for y in range(1, h + 1):
                for x in range(1, w + 1):
                    if not g[y][x] or g[y + dy][x + dx]:
                        continue  # 해당 방향이 비어있는 테두리 점만
                    nb = [g[y-1][x], g[y-1][x+1], g[y][x+1], g[y+1][x+1],
                          g[y+1][x], g[y+1][x-1], g[y][x-1], g[y-1][x-1]]
                    B = sum(nb)
                    if B < 2:
                        continue  # 획 끝점은 보존
                    A = sum(1 for i in range(8) if nb[i] == 0 and nb[(i + 1) % 8] == 1)
                    if A != 1:
                        continue  # 지우면 획이 끊김
                    # 1px 선 위의 점은 지우지 않음 (4방향 중 서로 마주보는 양쪽이 모두 비면 선)
                    if (nb[0] == 0 and nb[4] == 0) or (nb[2] == 0 and nb[6] == 0):
                        continue
                    g[y][x] = 0
                    changed = True
    return np.array(g, dtype=bool)[1:-1, 1:-1]


SS = 4  # 중심선을 뽑는 해상도 배율


def uniform_stroke_from_big(big, w, h, stroke):
    """고해상도 글자 → 중심선 추출 → 목표 크기로 축소 → 다시 1px로 → stroke 두께로 칠함.
    모든 획이 같은 픽셀 두께가 됨 (LED 비트맵 폰트처럼)."""
    import numpy as np
    # 획을 칠할 여유(stroke-1)만큼 작은 칸에서 중심선을 뽑아야 오른쪽·아래 끝이 잘리지 않음
    sw, sh_ = max(1, w - stroke + 1), max(1, h - stroke + 1)
    mid = big.resize((sw * SS, sh_ * SS), Image.BOX)
    a = np.asarray(mid.split()[3]) > 110
    sk = _thin(a)
    # 축소: SS×SS 칸에 중심선이 한 점이라도 지나가면 켬
    low = sk.reshape(sh_, SS, sw, SS).any(axis=(1, 3))
    low = _thin_seq(low)
    full = np.zeros((h, w), bool)
    for dy in range(stroke):
        for dx in range(stroke):
            full[dy:dy + sh_, dx:dx + sw] |= low
    low = full
    # 색: 원래 위치의 색(구분자 노랑 등) 유지, 빈 곳은 가까운 색으로 채움
    rgb = np.asarray(big.resize((w, h), Image.BOX).convert("RGBA"))[:, :, :3].copy()
    alpha = np.asarray(big.resize((w, h), Image.BOX))[:, :, 3]
    fill = rgb.copy()
    for _ in range(3):
        for sh in ((0, 1), (0, -1), (1, 0), (-1, 0)):
            nb = np.roll(fill, sh, (0, 1))
            nba = np.roll(alpha, sh, (0, 1))
            mask = (alpha < 40) & (nba >= 40)
            fill[mask] = nb[mask]
            alpha = np.where(mask, nba, alpha)
    out = np.zeros((h, w, 4), np.uint8)
    out[low, :3] = fill[low]
    out[low, 3] = 255
    return Image.fromarray(out, "RGBA")


# ───────────────────────── 내장 LED 숫자 폰트 ─────────────────────────
# 둥근 숫자: 각 숫자를 "중심선 경로"로 정의하고, s×s 붓으로 따라 그림 → 모든 획이 정확히 s픽셀.
# 좌표는 (0,0)~(X,Y) 상자 기준, c = 모서리 둥글기, M = 가운데 높이.
LED_CHARS = set("0123456789-")


def _digit_paths(ch, X, Y, c=None, s=1, k=None):
    if c is None:
        c = max(1, round(min(X, Y * 0.5) * 0.34))
    M = Y // 2
    R = lambda pts: [(X - x, Y - y) for x, y in pts]          # 180도 회전
    if ch == "0":
        return [[(c, 0), (X - c, 0), (X, c), (X, Y - c), (X - c, Y), (c, Y), (0, Y - c), (0, c), (c, 0)]]
    if ch == "1":
        sx = X - max(1, X // 4)
        return [[(max(0, sx - max(2, X * 2 // 3)), 0), (sx, 0), (sx, Y)]]   # 왼쪽 위 꼭지 + 기둥
    if ch == "2":
        k = max(1, Y // 6) if k is None else k
        return [[(0, c), (c, 0), (X - c, 0), (X, c), (X, M - c), (0, Y - k), (0, Y), (X, Y)]]
    if ch == "3":
        mx = max(c, X * 2 // 5)
        return [[(0, c), (c, 0), (X - c, 0), (X, c), (X, M - c), (X - c, M), (mx, M)],
                [(X - c, M), (X, M + c), (X, Y - c), (X - c, Y), (c, Y), (0, Y - c)]]
    if ch == "4":
        sx = max(X - max(1, X // 4), min(X, s + 1))   # 왼쪽 세로획과 최소 1px 띄움
        my = Y * 2 // 3
        return [[(sx, 0), (sx, Y)], [(0, 0), (0, my), (X, my)]]
    if ch == "5":
        return [[(X, 0), (0, 0), (0, M), (X - c, M), (X, M + c), (X, Y - c), (X - c, Y), (c, Y), (0, Y - c)]]
    if ch == "6":
        return [[(X, c), (X - c, 0), (c, 0), (0, c), (0, Y - c), (c, Y), (X - c, Y), (X, Y - c),
                 (X, M + c), (X - c, M), (0, M)]]
    if ch == "9":
        return [R(p) for p in _digit_paths("6", X, Y, c, s)]
    if ch == "7":
        k = max(1, Y // 6) if k is None else k
        return [[(0, 0), (X, 0), (X, k), (max(0, X * 2 // 5), Y)]]
    if ch == "8":
        return [[(c, 0), (X - c, 0), (X, c), (X, M - c), (X - c, M), (c, M), (0, M - c), (0, c), (c, 0)],
                [(c, M), (0, M + c), (0, Y - c), (c, Y), (X - c, Y), (X, Y - c), (X, M + c), (X - c, M)]]
    if ch == "-":
        return [[(0, M), (X, M)]]
    return []


def _line(x0, y0, x1, y1):
    pts = []
    dx, dy = abs(x1 - x0), -abs(y1 - y0)
    sx, sy = (1 if x0 < x1 else -1), (1 if y0 < y1 else -1)
    err = dx + dy
    while True:
        pts.append((x0, y0))
        if x0 == x1 and y0 == y1:
            return pts
        e2 = 2 * err
        if e2 >= dy:
            err += dy; x0 += sx
        if e2 <= dx:
            err += dx; y0 += sy


def _too_thick(g, s):
    """(s+1)×(s+1) 꽉 찬 덩어리가 있으면 True (= 어딘가 획이 s보다 두꺼움)"""
    import numpy as np
    k = s + 1
    if g.shape[0] < k or g.shape[1] < k:
        return False
    ii = np.pad(g.astype(np.int32), ((1, 0), (1, 0))).cumsum(0).cumsum(1)
    tot = ii[k:, k:] - ii[:-k, k:] - ii[k:, :-k] + ii[:-k, :-k]
    return bool((tot == k * k).any())


_glyph_cache = {}


def led_glyph(ch, w, h, s):
    """폭 w, 높이 h, 획 s 픽셀짜리 둥근 숫자 하나 (bool 배열).
    좁아서 모서리가 뭉치면 둥글기를 줄여서 다시 그림 → 항상 획 두께 s 보장."""
    import numpy as np
    key = (ch, w, h, s)
    if key in _glyph_cache:
        return _glyph_cache[key]
    X, Y = w - s, h - s
    c0 = max(1, round(min(X, Y * 0.5) * 0.34))
    # 둥글기(c)를 먼저 줄여보고, 그래도 뭉치면 꺾이는 세로 구간(k)을 늘려서 다시 그림
    tries = [(c, None) for c in range(c0, -1, -1)] + [(0, k) for k in range(max(1, Y // 6) + 1, Y // 2 + 1)]
    for c, k in tries:
        g = np.zeros((h, w), bool)
        for path in _digit_paths(ch, X, Y, c, s, k):
            for (x0, y0), (x1, y1) in zip(path, path[1:]):
                for x, y in _line(round(x0), round(y0), round(x1), round(y1)):
                    g[y:y + s, x:x + s] = True
        if not _too_thick(g, s):
            break
    _glyph_cache[key] = g
    return g


def led_widths(text, h, s, digit_w):
    return [max(s + 2, digit_w * 2 // 3) if ch == "-" else
            (max(s + 3, digit_w * 2 // 3) if ch == "1" else digit_w) for ch in text]


def render_led_digits(segs, target_h, max_w, stroke):
    """숫자/하이픈을 내장 LED 폰트로. 폭이 모자라면 숫자 폭→자간→획 순으로 줄임(한 줄 안에서는 항상 같은 획)."""
    import numpy as np
    text = "".join(t for t, _ in segs)
    s = stroke
    while True:
        gap = max(1, s // 2 + 1)
        for dw in range(round(target_h * 0.62), 2 * s + 1, -1):     # 숫자 폭 줄여보기 (안쪽 구멍 1px 이상 유지)
            ws = led_widths(text, target_h, s, dw)
            total = sum(ws) + gap * (len(text) - 1)
            if total <= max_w:
                break
        else:
            if gap > 1:
                gap = 1
                ws = led_widths(text, target_h, s, 2 * s + 1)
                total = sum(ws) + gap * (len(text) - 1)
            if total > max_w and s > 1:
                s -= 1
                continue
        break
    out = np.zeros((target_h, max(1, total), 4), np.uint8)
    colors = [c for t, c in segs for _ in t]
    x = 0
    for ch, cw, c in zip(text, ws, colors):
        g = led_glyph(ch, cw, target_h, s)
        sub = out[:, x:x + cw]
        sub[g, :3] = c
        sub[g, 3] = 255
        x += cw + gap
    return Image.fromarray(out[:, :max_w], "RGBA")


_pix_size_cache = {}


def pixel_size_for(font_path, target_h):
    """줄 높이(target_h)에 들어가는 가장 큰 정수 폰트 크기. 세로 확대·축소 없이 그대로 씀."""
    key = (font_path, target_h)
    if key not in _pix_size_cache:
        best = 6
        for size in range(6, target_h * 2 + 1):
            f = get_font(font_path, size)
            l, t, r, btm = f.getbbox("가Ag힣|", anchor="ls")
            if btm - t <= target_h:
                best = size
        _pix_size_cache[key] = best
    return _pix_size_cache[key]


def render_segments_pixel(segs, target_h, max_w, font_path):
    """픽셀(비트맵) 폰트용: 정수 크기로 안티앨리어싱 없이 그리고, 넘치면 가로만 압축."""
    size = pixel_size_for(font_path, target_h)
    font = get_font(font_path, size)
    total = int(sum(font.getlength(t) for t, _ in segs)) + 4
    base = size * 2
    img = Image.new("RGBA", (total, size * 3), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.fontmode = "1"
    x = 0.0
    for t, c in segs:
        d.text((x, base), t, font=font, fill=c + (255,), anchor="ls")
        x += font.getlength(t)
    bbox = img.getbbox()
    if not bbox:
        return Image.new("RGBA", (1, target_h), (0, 0, 0, 0))
    _, t, _, btm = font.getbbox("가Ag힣|", anchor="ls")
    top, bot = min(base + t, bbox[1]), max(base + btm, bbox[3])
    img = img.crop((bbox[0], top, bbox[2], bot))
    if img.width > max_w:
        img = img.resize((max_w, img.height), Image.NEAREST)
    out = Image.new("RGBA", (img.width, target_h), (0, 0, 0, 0))
    out.alpha_composite(img, (0, max(0, (target_h - img.height) // 2)))
    return out


def natural_width(text, target_h, font_path=None):
    font = get_font(font_path or text_font(), RENDER_SIZE)
    asc, desc = font.getmetrics()
    _, rt, _, rb = font.getbbox("힣갈Ag", anchor="la")
    return font.getlength(text) * target_h / (rb - rt)


def split_two(text):
    """두 줄로 나누기: 공백 중 가운데 가까운 곳, 없으면 글자 절반."""
    text = text.strip()
    spaces = [i for i, ch in enumerate(text) if ch == " "]
    if spaces:
        mid = len(text) / 2
        i = min(spaces, key=lambda s: abs(s - mid))
        return text[:i].strip(), text[i + 1:].strip()
    h = (len(text) + 1) // 2
    return text[:h], text[h:]


def paste_block(img, lines_imgs, box, align):
    """여러 줄 이미지를 box(x0,y0,x1,y1) 안에 세로 가운데 배치."""
    x0, y0, x1, y1 = box
    gap = 1
    total_h = sum(i.height for i in lines_imgs) + gap * (len(lines_imgs) - 1)
    y = y0 + (y1 - y0 - total_h) // 2
    for li in lines_imgs:
        if align == "왼쪽":
            x = x0
        else:
            x = x0 + (x1 - x0 - li.width) // 2
        img.alpha_composite(li, (x, y))
        y += li.height + gap


def fit_text(text, color, box, align, one_line_h, two_line_h):
    """한 줄로 충분하면 크게, 아니면 두 줄."""
    x0, y0, x1, y1 = box
    bw = x1 - x0
    text = (text or "").strip()
    if not text:
        return []
    if natural_width(text, one_line_h) <= bw * 1.5 or " " not in text and len(text) <= 3:
        return [render_segments([(text, color)], one_line_h, bw)]
    a, b = split_two(text)
    return [render_segments([(a, color)], two_line_h, bw),
            render_segments([(b, color)], two_line_h, bw)]


def draw_arrow(img, cx, cy, shape, color):
    d = ImageDraw.Draw(img)
    c = color + (255,)
    if shape == "◆":
        # 원본과 동일한 도트: 가로 15 × 세로 13 마름모, 가운데 세로줄은 위·아래 4칸만 비움(가운데 5칸은 채움)
        for dy in range(-6, 7):
            for dx in range(-(7 - abs(dy)), 7 - abs(dy) + 1):
                if dx == 0 and abs(dy) >= 3:
                    continue
                img.putpixel((cx + dx, cy + dy), c)
    elif shape == "▶":
        d.polygon([(cx - 4, cy - 6), (cx + 5, cy), (cx - 4, cy + 6)], fill=c)
    elif shape == "→":
        d.line([(cx - 6, cy), (cx + 5, cy)], fill=c, width=2)
        d.polygon([(cx + 6, cy), (cx + 1, cy - 4), (cx + 1, cy + 4)], fill=c)
    elif shape == "↔":
        d.line([(cx - 6, cy), (cx + 6, cy)], fill=c, width=2)
        d.polygon([(cx + 7, cy), (cx + 3, cy - 4), (cx + 3, cy + 4)], fill=c)
        d.polygon([(cx - 7, cy), (cx - 3, cy - 4), (cx - 3, cy + 4)], fill=c)


def has_hangul(s):
    return any("\uac00" <= ch <= "\ud7a3" or "\u3131" <= ch <= "\u318e" for ch in s)


def parse_via(via):
    """경유지 입력 → (한글 줄들, 영문 줄들). 줄바꿈=줄, 쉼표=정류장."""
    ko, en = [], []
    for raw in (via or "").replace("\r", "").split("\n"):
        stops = [s.strip() for s in raw.split(",") if s.strip()]
        if not stops:
            continue
        (ko if has_hangul(raw) else en).append(stops)
    return ko[:2], en[:2]


def via_segments(stops, color, sep_color):
    segs = []
    for i, s in enumerate(stops):
        if i:
            segs.append(("-", sep_color))
        segs.append((s, color))
    return segs


# 템플릿: 각 템플릿은 렌더 함수
def tpl_jinmyeong_uniform(r, st):
    """(이전 방식) 획 두께 완전 균일형"""
    img = Image.new("RGBA", (W, H), (0, 0, 0, 255))
    num = str(r.get("no", "")).strip()
    align = st["align"]
    yel_big, yel_small = col(st["num_big"]), col(st["num_small"])
    band = 32

    # 1) 큰 번호
    if num:
        n = render_segments([(num, yel_big)], 24, W - 8, num_font())
        img.alpha_composite(n, ((W - n.width) // 2, (band - n.height) // 2 + 1))

    NUM_W = 34
    cx0 = NUM_W + 3

    def small_num(y):
        if num:
            n = render_segments([(num, yel_small)], 20, NUM_W, num_font())
            img.alpha_composite(n, ((NUM_W - n.width) // 2 + 1, y + (band - n.height) // 2))

    def od_band(y, start, end):
        small_num(y)
        mid = cx0 + (W - cx0) * 45 // 100   # = 106 (원본 화살표 중심)
        left = (cx0, y + 1, mid - 8, y + band - 1)
        right = (mid + 11, y + 1, W - 1, y + band - 1)
        paste_block(img, fit_text(start, col(st["start"]), left, align, 22, 14), left, align)
        paste_block(img, fit_text(end, col(st["end"]), right, align, 22, 14), right, align)
        if start or end:
            draw_arrow(img, mid, y + band // 2, st["arrow_shape"], col(st["arrow"]))

    # 2) 한글 기점-종점, 3) 영문 기점-종점
    od_band(band * 1, r.get("start_ko", ""), r.get("end_ko", ""))
    od_band(band * 2, r.get("start_en", ""), r.get("end_en", ""))

    # 4) 한글 경유지, 5) 영문 경유지
    ko, en = parse_via(r.get("via", ""))
    for i, lines in enumerate((ko, en)):
        y = band * (3 + i)
        small_num(y)
        box = (cx0, y + 1, W - 1, y + band - 1)
        h = 22 if len(lines) == 1 else 14
        imgs = [render_segments(via_segments(s, col(st["via"]), col(st["sep"])), h, box[2] - box[0])
                for s in lines]
        if imgs:
            paste_block(img, imgs, box, align)
    return img.convert("RGB")


# ───────────────────────── 경고 · 글자 보정표 ─────────────────────────
WARN = []          # 렌더 중 생긴 압축 경고 [(내용, 비율)]
GLYPHS = {}        # 글자 보정표 {"글꼴|줄높이|글자": ["..##..", ...]}
GLYPH_FILE = APP_DIR / "glyphs.json"
WARN_RATIO = 0.90  # 이보다 많이 줄어들면 경고


def _warn(text, ratio, limit=None):
    """limit 보다 많이 줄면 경고. 같은 내용은 한 번만."""
    if ratio < (WARN_RATIO if limit is None else limit) and (text, ratio) not in WARN:
        WARN.append((text, ratio))


def load_glyphs():
    GLYPHS.clear()
    try:
        GLYPHS.update(json.loads(GLYPH_FILE.read_text(encoding="utf-8")))
    except Exception:
        pass
    return GLYPHS


def save_glyphs():
    APP_DIR.mkdir(parents=True, exist_ok=True)
    GLYPH_FILE.write_text(json.dumps(GLYPHS, ensure_ascii=False, indent=1), encoding="utf-8")


def glyph_key(kind, h, ch):
    """kind: 'TEXT' 또는 'NUM'. 글자 글꼴이 바뀌면 보정표도 따로 관리됨."""
    font = Path(text_font() or "").name if kind == "TEXT" else "NUM"
    return f"{font}|{h}|{ch}"


def glyph_override(kind, h, ch):
    rows = GLYPHS.get(glyph_key(kind, h, ch))
    if not rows:
        return None
    import numpy as np
    return np.array([[c == "#" for c in r] for r in rows], bool)


def bool_to_rgba(m, color):
    import numpy as np
    rgba = np.zeros(m.shape + (4,), np.uint8)
    rgba[m, :3] = color
    rgba[m, 3] = 255
    return Image.fromarray(rgba, "RGBA")


def default_glyph(kind, h, ch, room, kx=0.96):
    """보정 전 기본 모양 (글자 보정 창에서 출발점으로 씀)"""
    import numpy as np
    saved = dict(GLYPHS)
    GLYPHS.clear()
    try:
        if kind == "NUM":
            im = render_num_orig(ch, (255, 255, 255), h, 10 ** 4, kx=kx, th=100, track=0)
        else:
            im = render_segments([(ch, (255, 255, 255))], h, room)
    finally:
        GLYPHS.update(saved)
    return np.asarray(im)[:, :, 3] > 0


# ───────────────────────── 원본 보정형 렌더러 ─────────────────────────
# 원본(EBASE) 화면을 점 단위로 측정해서 맞춘 값들
ORIG_NUM_FONT = BASE_DIR / "fonts" / "LiberationSans-Bold.ttf"   # Arial Bold와 같은 모양의 무료 글꼴
_NUM_SS = 8
_num_em_cache = {}


def _num_em(path, target_h):
    """숫자 잉크 높이가 target_h가 되는 글꼴 크기"""
    key = (path, target_h)
    if key not in _num_em_cache:
        best = None
        for em10 in range(int(target_h * 10), int(target_h * 20)):
            em = em10 / 10
            f = get_font(path, int(round(em * _NUM_SS)))
            l, t, r, bt = f.getbbox("0123456789", anchor="ls")
            hh = round((bt - t) / _NUM_SS)
            if hh == target_h:
                best = em
                break
        _num_em_cache[key] = best or target_h * 1.4
    return _num_em_cache[key]


def squeeze_w(img, w):
    """가로로 넘치면 세로줄을 빼서 맞춤. 원본은 아무 줄이나 빼서 글자가 깨지지만('Gangdorg'),
    여기서는 ① 빈 줄 ② 빼도 획이 붙거나 끊기지 않는 줄 순서로 골라 빼서 글자 모양과 획 두께를 지킴."""
    import numpy as np
    if img.width <= w:
        return img
    a = np.asarray(img)
    on = a[:, :, 3] > 0
    keep = list(range(a.shape[1]))
    need = a.shape[1] - w
    def cost(i):
        c = on[:, keep[i]]
        if not c.any():
            return 0
        L = on[:, keep[i - 1]] if i > 0 else np.zeros_like(c)
        R = on[:, keep[i + 1]] if i + 1 < len(keep) else np.zeros_like(c)
        # 빼면 좌우 열이 바로 붙음: 같은 행에 둘 다 켜져 있으면 2px 덩어리 위험
        merge = int((L & R).sum()) + int(((L[:-1] & R[1:]) | (L[1:] & R[:-1])).sum())
        return 1 + c.sum() + 50 * merge
    while need > 0:
        # 균등하게 퍼지도록 구간마다 가장 싼 열을 하나씩 뺌 (한 번에 한 구간씩 처리)
        seg = max(2, len(keep) // need)
        drop = []
        for start in range(0, len(keep), seg):
            if len(drop) >= need:
                break
            drop.append(min(range(start, min(start + seg, len(keep))), key=cost))
        drop = set(drop[:need])
        keep = [k for i, k in enumerate(keep) if i not in drop]
        need = len(keep) - w
    return Image.fromarray(a[:, keep], "RGBA")


def render_num_orig(text, color, target_h, max_w, kx=1.0, th=80, track=1):
    """원본 번호 방식: Arial Bold형 글꼴을 비AA로 찍고, 글자마다 정수 위치, 넘치면 가로 압축."""
    import numpy as np
    path = str(ORIG_NUM_FONT) if ORIG_NUM_FONT.exists() else num_font()
    em = _num_em(path, target_h)
    f = get_font(path, int(round(em * _NUM_SS)))
    asc, desc = f.getmetrics()
    _, rt, _, rb = f.getbbox("0123456789", anchor="ls")
    cols = []
    for ch in text:
        if ch == " ":
            cols.append(None); continue
        adv = f.getlength(ch)
        ov = glyph_override("NUM", target_h, ch)
        if ov is not None and ov.shape[0] == target_h:
            cols.append((ov, max(ov.shape[1], round(adv * kx / _NUM_SS))))
            continue
        big = Image.new("L", (int(adv + em * _NUM_SS), asc + desc), 0)
        ImageDraw.Draw(big).text((0, asc), ch, font=f, fill=255, anchor="ls")
        big = big.crop((0, asc + rt, big.width, asc + rb))
        sm = big.resize((max(1, round(big.width * kx / _NUM_SS)), target_h), Image.BOX)
        cols.append((np.asarray(sm) > th, round(adv * kx / _NUM_SS)))
    total = sum((c[1] if c else round(em * 0.28)) + track for c in cols) + target_h
    out = np.zeros((target_h, max(1, total)), bool)
    x = 0
    for c in cols:
        if c is None:
            x += round(em * 0.28) + track; continue
        m, adv = c
        w = min(m.shape[1], out.shape[1] - x)
        out[:, x:x + w] |= m[:, :w]
        x += adv + track
    xs = np.where(out.any(0))[0]
    if len(xs) == 0:
        return Image.new("RGBA", (1, target_h), (0, 0, 0, 0))
    out = out[:, xs.min():xs.max() + 1]
    rgba = np.zeros(out.shape + (4,), np.uint8)
    rgba[out, :3] = color
    rgba[out, 3] = 255
    im = Image.fromarray(rgba, "RGBA")
    if im.width > max_w:       # 번호는 원본과 같은 최근접 압축이 가장 일치함
        _warn(f"번호 {text}", max_w / im.width, limit=0.62)   # 작은 번호는 원본도 75% 정도로 눌러 그림
        im = im.resize((max_w, im.height), Image.NEAREST)
    return im


def is_wide(ch):
    o = ord(ch)
    return (0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F
            or 0x4E00 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF)


def render_mono(segs, target_h, full, half, max_w):
    """원본처럼 고정폭 배치: 한글 = full칸, 영문·숫자·기호 = half칸.
    줄이 넘치면 칸을 비율대로 좁히고 글자를 그 폭에 맞춰 '다시 그림' → 좁아져도 획은 1픽셀 유지."""
    chars = [(ch, color) for text, color in segs for ch in text]
    widths = [full if is_wide(ch) else half for ch, _ in chars]
    total = sum(widths)
    k = min(1.0, max_w / total) if total else 1.0
    if k < 1.0:
        _warn("".join(ch for ch, _ in chars), k)
    out = Image.new("RGBA", (max(1, int(total * k) + 2), target_h), (0, 0, 0, 0))
    x = 0.0
    for (ch, color), cw in zip(chars, widths):
        cell = cw * k
        x0 = int(round(x))
        room = max(1, int(round(x + cell)) - x0 - 1)      # 옆 글자와 최소 1칸 띄움
        if ch == "-":
            # 원본 하이픈: 칸을 꽉 채우는 가로줄 (칸 7→6px, 6→6px, 압축 시 칸 폭만큼), 높이는 줄 가운데
            ln = max(1, int(cell) - (1 if cell >= 6.5 else 0))
            for xx in range(x0, min(out.width, x0 + ln)):
                out.putpixel((xx, target_h // 2), tuple(color) + (255,))
        elif ch != " ":
            ov = glyph_override("TEXT", target_h, ch)
            if ov is not None and ov.shape[0] == target_h:
                g = squeeze_w(bool_to_rgba(ov, color), room)
            else:
                g = render_segments([(ch, color)], target_h, room)
            if g.getbbox():              # 폭 1픽셀짜리 글자(l, i, I)도 그림
                out.alpha_composite(g, (x0 + (room - g.width) // 2, 0))
        x += cell
    bx = out.getbbox()
    if bx:
        out = out.crop((bx[0], 0, bx[2], target_h))
    return squeeze_w(out, max_w)


# 줄 규격 (원본 측정값): (줄높이, 한글 칸폭, 반각 칸폭)
SPEC_KO_TWO = (14, 13, 6)    # 한글 두 줄 기점·종점
SPEC_EN_TWO = (14, 13, 7)    # 영문 두 줄 기점·종점
SPEC_BIG_KO = (24, 23, 12)   # 한 줄 한글 종점 ("강일IC")
SPEC_EN_ONE = (13, 13, 7)    # 한 줄 영문 ("Gang-il IC")
SPEC_VIA_KO = (13, 13, 6)    # 한글 경유지
SPEC_VIA_EN = (13, 13, 7)    # 영문 경유지


def tpl_jinmyeong_front(r, st):
    """진명 앞판 192×160 — 원본 화면에 맞춰 보정한 배치"""
    img = Image.new("RGBA", (W, H), (0, 0, 0, 255))
    num = str(r.get("no", "")).strip()
    align = st["align"]
    band = 32
    use_orig_num = CFG["led_digits"] and st.get("num_weight") == "원본형"   # 원본과 같은 Arial Bold형 숫자

    kind_rule = kind_style(r.get("kind"), st)

    def number(color, h, max_w, kx, track):
        if kind_rule:
            color = kind_rule["color"]
        if use_orig_num:
            im = render_num_orig(num, color, h, max_w, kx=kx, th=100, track=track)
        else:
            im = render_segments([(num, color)], h, max_w, num_font())
        return invert_box(im, color) if kind_rule and kind_rule["invert"] else im

    # 큰 번호: 높이 22, 위에서 5
    if num:
        n = number(col(st["num_big"]), 22, W - 8, 0.96, 1)
        img.alpha_composite(n, ((W - n.width) // 2, 5 - (n.height - 22) // 2))

    def small_num(y):
        if num:
            n = number(col(st["num_small"]), 19, 30, 0.96, 0)
            img.alpha_composite(n, (max(0, 1 + (30 - n.width) // 2), y + 6 - (n.height - 19) // 2))

    def put(lines, box, tops):
        x0, x1 = box
        for li, top in zip(lines, tops):
            x = x0 if align == "왼쪽" else x0 + (x1 - x0 - li.width) // 2
            img.alpha_composite(li, (x, top))

    def place_text(text, color, box, y, korean):
        text = (text or "").strip()
        if not text:
            return
        bw = box[1] - box[0]
        if korean:
            h, fw, hw = SPEC_BIG_KO
            one = render_mono([(text, color)], h, fw, hw, 10 ** 4)
            if one.width <= bw:
                put([one], box, [y + 4]); return
        else:
            h, fw, hw = SPEC_EN_ONE
            one = render_mono([(text, color)], h, fw, hw, 10 ** 4)
            if one.width <= bw:
                put([one], box, [y + 9]); return
        a, b2 = split_two(text)
        h, fw, hw = SPEC_KO_TWO if korean else SPEC_EN_TWO
        put([render_mono([(a, color)], h, fw, hw, bw), render_mono([(b2, color)], h, fw, hw, bw)], box, [y + 2, y + 17])

    def od_band(y, start, end, korean):
        small_num(y)
        place_text(start, col(st["start"]), (36, 91), y, korean)
        place_text(end, col(st["end"]), (115, 192), y, korean)
        if start or end:
            draw_arrow(img, 106, y + 16, st["arrow_shape"], col(st["arrow"]))

    od_band(band * 1, r.get("start_ko", ""), r.get("end_ko", ""), True)
    od_band(band * 2, r.get("start_en", ""), r.get("end_en", ""), False)

    # 경유지: 한글 줄 위치 +3/+18, 영문 +1/+16, 한 줄이면 가운데
    ko, en = parse_via(r.get("via", ""))
    for i, lines in enumerate((ko, en)):
        y = band * (3 + i)
        small_num(y)
        if not lines:
            continue
        h, fw, hw = SPEC_VIA_KO if i == 0 else SPEC_VIA_EN
        box = (36, 189)
        imgs = [render_mono(via_segments(s, col(st["via"]), col(st["sep"])), h, fw, hw, box[1] - box[0]) for s in lines]
        tops = [y + 9] if len(imgs) == 1 else ([y + 3, y + 18] if i == 0 else [y + 1, y + 16])
        put(imgs, box, tops)
    return img.convert("RGB")


def tpl_simple_front(r, st):
    """번호 + 종점만 크게 (간단형)"""
    img = Image.new("RGBA", (W, H), (0, 0, 0, 255))
    num = str(r.get("no", "")).strip()
    n = render_segments([(num, col(st["num_big"]))], 60, W - 8, num_font())
    img.alpha_composite(n, ((W - n.width) // 2, 10))
    box = (4, 84, W - 4, 124)
    paste_block(img, fit_text(r.get("end_ko", ""), col(st["end"]), box, "가운데", 34, 18), box, "가운데")
    box = (4, 126, W - 4, 156)
    paste_block(img, fit_text(r.get("end_en", ""), col(st["end"]), box, "가운데", 22, 13), box, "가운데")
    return img.convert("RGB")


TEMPLATES = {
    "jinmyeong_front_192x160": tpl_jinmyeong_front,
    "jinmyeong_uniform_192x160": tpl_jinmyeong_uniform,
    "simple_front_192x160": tpl_simple_front,
}


def render_route(route, style=None):
    fn = TEMPLATES.get(route.get("template") or "", None) or tpl_jinmyeong_front
    st = resolve_style(style)
    tf, nf = st["text_font"], st["num_font"]
    CFG["text"] = tf if tf and os.path.exists(tf) else TEXT_FONT
    CFG["num"] = nf if nf and os.path.exists(nf) else NUM_FONT
    CFG["pixel"] = str(st["pixel_mode"]) in ("1", "True", "true")
    CFG["stroke"] = WEIGHTS.get(st["weight"], 1)
    CFG["num_stroke"] = NUM_WEIGHTS.get(st["num_weight"], 3)
    CFG["led_digits"] = not nf      # 번호 폰트 (자동) = 내장 LED 숫자
    return fn(route, st)


def out_name(route):
    no = str(route.get("no", "noname")).strip() or "noname"
    safe = "".join(ch for ch in no if ch not in '\\/:*?"<>|')
    return f"{safe}_{safe}.bmp"



# ───────────────────────── 노선 구분 (급행·심야 등) ─────────────────────────
KINDS = ["일반", "급행", "좌석", "심야", "마을", "순환"]
KIND_DEFAULTS = {  # 구분: (번호색, 반전)
    "일반": ("(기본)", False), "급행": ("빨강", True), "좌석": ("빨강", False),
    "심야": ("하늘", True), "마을": ("초록", False), "순환": ("주황", False),
}


def kind_style(kind, st):
    """구분별 번호 표시 규칙. 일반/빈칸이면 None (기존과 완전히 동일)"""
    if not kind or kind == "일반":
        return None
    color, inv = KIND_DEFAULTS.get(kind, ("(기본)", False))
    custom = (st.get("kinds") or {}).get(kind)
    if custom:
        color, inv = custom.get("color", color), bool(custom.get("invert", inv))
    if color in (None, "", "(기본)") and not inv:
        return None
    rgb = COLORS.get(color) or col(st.get("num_big", "노랑"))
    return {"color": rgb, "invert": inv}


def invert_box(im, color):
    """반전 표시: 색칠한 상자 안에 검은 숫자 (상자는 글자보다 사방 1칸 큼)"""
    import numpy as np
    a = np.asarray(im)[:, :, 3] > 0
    h, w = a.shape
    box = np.ones((h + 2, w + 4), bool)
    box[1:h + 1, 2:w + 2] &= ~a
    return bool_to_rgba(box, color)


# ───────────────────────── 점 편집 (노선별) ─────────────────────────
def apply_edits(img, edits):
    """edits: {"x,y": [r,g,b]} — [0,0,0]은 끄기"""
    if not edits:
        return img
    img = img.copy()
    W_, H_ = img.size
    for key, rgb in edits.items():
        try:
            x, y = map(int, key.split(","))
        except ValueError:
            continue
        if 0 <= x < W_ and 0 <= y < H_:
            img.putpixel((x, y), tuple(int(v) for v in rgb[:3]))
    return img


# ───────────────────────── 설정 파일 템플릿 (크기·종류 추가) ─────────────────────────
TEMPLATE_DIR = APP_DIR / "templates"
TEXT_FIELDS = {"start_ko": "기점(한)", "end_ko": "종점(한)", "start_en": "기점(영)", "end_en": "종점(영)",
               "via_ko": "한글 경유지(첫 줄)", "via_en": "영문 경유지(첫 줄)"}
# 줄 규격 이름 → (줄높이, 한글칸, 반각칸)
SPECS = {"big_ko": SPEC_BIG_KO, "ko_two": SPEC_KO_TWO, "en_two": SPEC_EN_TWO, "en_one": SPEC_EN_ONE,
         "via_ko": SPEC_VIA_KO, "via_en": SPEC_VIA_EN, "small": (11, 11, 6), "tiny": (9, 9, 5)}

BUILTIN_JSON_TEMPLATES = {
    "side_160x32 (측면, 한/영 2장)": {
        "width": 160, "height": 32,
        "pages": [
            {"name": "한글", "elements": [
                {"type": "number", "x": 0, "y": 0, "w": 44, "h": 32, "digit_h": 22},
                {"type": "text", "text": "{end_ko}", "x": 46, "y": 0, "w": 114, "h": 32,
                 "one": "big_ko", "two": "ko_two", "color": "end"}]},
            {"name": "영문", "elements": [
                {"type": "number", "x": 0, "y": 0, "w": 44, "h": 32, "digit_h": 22},
                {"type": "text", "text": "{end_en}", "x": 46, "y": 0, "w": 114, "h": 32,
                 "one": "en_one", "two": "en_two", "color": "end"}]},
        ]},
    "front_192x32 (전면 한 줄, 한/영 2장)": {
        "width": 192, "height": 32,
        "pages": [
            {"name": "한글", "elements": [
                {"type": "number", "x": 0, "y": 0, "w": 36, "h": 32, "digit_h": 19},
                {"type": "text", "text": "{start_ko}", "x": 36, "y": 0, "w": 55, "h": 32, "one": "big_ko", "two": "ko_two", "color": "start"},
                {"type": "arrow", "x": 106, "y": 16},
                {"type": "text", "text": "{end_ko}", "x": 115, "y": 0, "w": 77, "h": 32, "one": "big_ko", "two": "ko_two", "color": "end"}]},
            {"name": "영문", "elements": [
                {"type": "number", "x": 0, "y": 0, "w": 36, "h": 32, "digit_h": 19},
                {"type": "text", "text": "{start_en}", "x": 36, "y": 0, "w": 55, "h": 32, "one": "en_one", "two": "en_two", "color": "start"},
                {"type": "arrow", "x": 106, "y": 16},
                {"type": "text", "text": "{end_en}", "x": 115, "y": 0, "w": 77, "h": 32, "one": "en_one", "two": "en_two", "color": "end"}]},
        ]},
    "rear_64x32 (후면 번호)": {
        "width": 64, "height": 32,
        "pages": [{"name": "번호", "elements": [
            {"type": "number", "x": 0, "y": 0, "w": 64, "h": 32, "digit_h": 24}]}]},
}


def load_json_templates():
    """기본 템플릿 + ~/BusSignMaker/templates/*.json (사용자가 추가한 판)"""
    out = dict(BUILTIN_JSON_TEMPLATES)
    try:
        for p in sorted(TEMPLATE_DIR.glob("*.json")):
            try:
                t = json.loads(p.read_text(encoding="utf-8"))
                out[t.get("name") or p.stem] = t
            except Exception as e:  # 잘못된 파일은 건너뜀
                WARN.append((f"템플릿 {p.name} 읽기 실패: {e}", 0))
    except Exception:
        pass
    return out


def write_example_templates():
    """사용자가 고쳐 쓸 수 있게 예시 템플릿 파일을 만들어 둠 (이미 있으면 그대로)"""
    try:
        TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)
        ex = TEMPLATE_DIR / "_예시_side_160x32.json.txt"
        if not ex.exists():
            t = dict(BUILTIN_JSON_TEMPLATES["side_160x32 (측면, 한/영 2장)"])
            t["name"] = "내 측면판 160x32"
            ex.write_text("// 이 파일을 복사해서 확장자를 .json 으로 바꾸면 템플릿 목록에 나타납니다.\n"
                          "// type: number / text / arrow,  one·two: " + ", ".join(SPECS) + "\n"
                          + json.dumps(t, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def _fill(text, route):
    ko, en = parse_via(route.get("via", ""))
    vals = dict(route)
    vals["via_ko"] = "-".join(ko[0]) if ko else ""
    vals["via_en"] = "-".join(en[0]) if en else ""
    try:
        return text.format_map({k: (v if isinstance(v, str) else str(v)) for k, v in vals.items()
                                if not isinstance(v, (dict, list))} | _Missing())
    except Exception:
        return text


class _Missing(dict):
    def __missing__(self, k):
        return ""


def render_json_page(tpl, page, route, st):
    w, h = int(tpl["width"]), int(tpl["height"])
    img = Image.new("RGBA", (w, h), (0, 0, 0, 255))
    num = str(route.get("no", "")).strip()
    rule = kind_style(route.get("kind"), st)
    for el in page.get("elements", []):
        t = el.get("type")
        x, y = int(el.get("x", 0)), int(el.get("y", 0))
        bw, bh = int(el.get("w", w - x)), int(el.get("h", h - y))
        if t == "number" and num:
            dh = int(el.get("digit_h", max(8, bh - 8)))
            color = rule["color"] if rule else col(st[el.get("color", "num_big")] if el.get("color", "num_big") in st else "노랑")
            n = render_num_orig(num, color, dh, bw - (4 if rule and rule["invert"] else 0), kx=float(el.get("kx", 0.96)), th=100, track=1)
            if rule and rule["invert"]:
                n = invert_box(n, color)
            img.alpha_composite(n, (x + (bw - n.width) // 2, y + (bh - n.height) // 2))
        elif t == "arrow":
            draw_arrow(img, x, y, el.get("shape", st["arrow_shape"]), col(st["arrow"]))
        elif t == "text":
            text = _fill(el.get("text", ""), route).strip()
            if not text:
                continue
            color = col(st.get(el.get("color", "end"), "흰색"))
            h1, f1, s1 = SPECS.get(el.get("one", "big_ko"), SPEC_BIG_KO)
            one = render_mono([(text, color)], h1, f1, s1, 10 ** 4)
            if one.width <= bw or not el.get("two"):
                one = squeeze_w(one, bw) if one.width > bw else one
                if one.width < 10 ** 4 and one.width > bw:
                    pass
                if render_mono([(text, color)], h1, f1, s1, 10 ** 4).width > bw:
                    _warn(text, bw / max(1, render_mono([(text, color)], h1, f1, s1, 10 ** 4).width))
                img.alpha_composite(one, (x + (bw - one.width) // 2, y + (bh - h1) // 2))
            else:
                h2, f2, s2 = SPECS.get(el.get("two"), SPEC_KO_TWO)
                a, b2 = split_two(text)
                gap = max(1, (bh - 2 * h2) // 3)
                for i, part in enumerate((a, b2)):
                    li = render_mono([(part, color)], h2, f2, s2, bw)
                    img.alpha_composite(li, (x + (bw - li.width) // 2, y + gap + i * (h2 + gap)))
    return img.convert("RGB")


def render_pages(route, style=None, with_edits=True):
    """노선 → [페이지 이미지...]. 점 편집까지 적용. WARN 에 압축 경고가 쌓임."""
    WARN.clear()
    tname = route.get("template") or ""
    jt = load_json_templates().get(tname)
    if jt is None:
        pages = [render_route(route, style)]
    else:
        render_route({"no": ""}, style)          # 글꼴·획 설정만 적용
        st = resolve_style(style)
        pages = [render_json_page(jt, pg, route, st) for pg in jt.get("pages", [])] or [Image.new("RGB", (jt["width"], jt["height"]))]
    if with_edits:
        ed = route.get("edits") or {}
        pages = [apply_edits(p, ed.get(str(i))) for i, p in enumerate(pages)]
    return pages


def all_template_names():
    return list(TEMPLATES) + list(load_json_templates())


# ───────────────────────── 한글 → 영문 (국어의 로마자 표기법) ─────────────────────────
_CHO = ["g", "kk", "n", "d", "tt", "r", "m", "b", "pp", "s", "ss", "", "j", "jj", "ch", "k", "t", "p", "h"]
_JUNG = ["a", "ae", "ya", "yae", "eo", "e", "yeo", "ye", "o", "wa", "wae", "oe", "yo", "u", "wo", "we", "wi",
         "yu", "eu", "ui", "i"]
# 받침 번호(0=없음) → (대표음, 뒤에 모음이 오면 넘어가는 소리(초성 번호), 남는 받침 대표음)
_JONG_NAMES = ["", "ㄱ", "ㄲ", "ㄳ", "ㄴ", "ㄵ", "ㄶ", "ㄷ", "ㄹ", "ㄺ", "ㄻ", "ㄼ", "ㄽ", "ㄾ", "ㄿ", "ㅀ", "ㅁ",
               "ㅂ", "ㅄ", "ㅅ", "ㅆ", "ㅇ", "ㅈ", "ㅊ", "ㅋ", "ㅌ", "ㅍ", "ㅎ"]
_JONG_SOUND = {"": "", "ㄱ": "k", "ㄲ": "k", "ㄳ": "k", "ㄴ": "n", "ㄵ": "n", "ㄶ": "n", "ㄷ": "t", "ㄹ": "l",
               "ㄺ": "k", "ㄻ": "m", "ㄼ": "l", "ㄽ": "l", "ㄾ": "l", "ㄿ": "p", "ㅀ": "l", "ㅁ": "m", "ㅂ": "p",
               "ㅄ": "p", "ㅅ": "t", "ㅆ": "t", "ㅇ": "ng", "ㅈ": "t", "ㅊ": "t", "ㅋ": "k", "ㅌ": "t", "ㅍ": "p", "ㅎ": "t"}
# 뒤 음절이 ㅇ으로 시작할 때 (앞에 남는 소리, 넘어가는 소리)
_LIAISON = {"ㄱ": ("", "g"), "ㄲ": ("", "kk"), "ㄳ": ("k", "s"), "ㄴ": ("", "n"), "ㄵ": ("n", "j"), "ㄶ": ("", "n"),
            "ㄷ": ("", "d"), "ㄹ": ("", "r"), "ㄺ": ("l", "g"), "ㄻ": ("l", "m"), "ㄼ": ("l", "b"), "ㄽ": ("l", "s"),
            "ㄾ": ("l", "t"), "ㄿ": ("l", "p"), "ㅀ": ("", "r"), "ㅁ": ("", "m"), "ㅂ": ("", "b"), "ㅄ": ("p", "s"),
            "ㅅ": ("", "s"), "ㅆ": ("", "ss"), "ㅈ": ("", "j"), "ㅊ": ("", "ch"), "ㅋ": ("", "k"), "ㅌ": ("", "t"),
            "ㅍ": ("", "p"), "ㅎ": ("", "")}


def _decompose(ch):
    o = ord(ch) - 0xAC00
    return o // 588, (o % 588) // 28, _JONG_NAMES[o % 28]


def romanize_word(word):
    """한글 단어 하나 → 로마자 (연음·비음화·유음화·구개음화·ㅎ축약 반영. 명사에서 ㄱㄷㅂ+ㅎ은 h를 그대로 둠)"""
    syl = [_decompose(c) for c in word]
    n = len(syl)
    onset = [_CHO[c] for c, _, _ in syl]
    coda = [_JONG_SOUND[j] for _, _, j in syl]
    sep = [""] * n
    for i in range(n - 1):
        J = syl[i][2]
        ncho, nvow = syl[i + 1][0], syl[i + 1][1]
        if J == "":
            continue
        if ncho == 11:                                   # 뒤가 ㅇ → 연음
            if J == "ㅇ":
                sep[i + 1] = "-"                         # 강일 → Gang-il (읽기 헷갈림 방지)
                continue
            keep, move = _LIAISON[J]
            if J in ("ㄷ", "ㅌ") and nvow == 20:         # 구개음화: 같이 → gachi
                move = "j" if J == "ㄷ" else "ch"
            coda[i], onset[i + 1] = keep, move
        elif J in ("ㅎ", "ㄶ", "ㅀ"):
            coda[i] = {"ㅎ": "", "ㄶ": "n", "ㅀ": "l"}[J]
            if ncho in (0, 3, 12):                       # 좋고 → joko
                onset[i + 1] = {0: "k", 3: "t", 12: "ch"}[ncho]
        elif ncho in (2, 6):                             # 뒤가 ㄴ·ㅁ → 비음화
            if coda[i] == "l" and ncho == 2:
                onset[i + 1] = "l"                       # 설날 → seollal
            else:
                coda[i] = {"k": "ng", "t": "n", "p": "m"}.get(coda[i], coda[i])
        elif ncho == 5:                                  # 뒤가 ㄹ
            if coda[i] in ("n", "l"):
                coda[i], onset[i + 1] = "l", "l"         # 신라 → Silla
            elif coda[i] in ("m", "ng"):
                onset[i + 1] = "n"                       # 종로 → Jongno
            else:
                coda[i], onset[i + 1] = {"k": "ng", "t": "n", "p": "m"}[coda[i]], "n"   # 독립 → dongnip
    return "".join(sep[i] + onset[i] + _JUNG[syl[i][1]] + coda[i] for i in range(n))


ROMAN_FILE = APP_DIR / "romanize_dict.json"
ROMAN_DEFAULT = {
    "_설명": "exact = 이름 그대로 바꿈, suffix = 끝말 바꿈 (긴 것부터 적용). 자유롭게 고치세요.",
    "exact": {"서울역": "Seoul Station", "부산역": "Busan Station", "서면": "Seomyeon"},
    "suffix": {"환승센터": " Transfer Center", "버스터미널": " Bus Terminal", "터미널": " Terminal",
               "역": " Stn", "입구": " Entrance", "사거리": " Intersection", "삼거리": " Junction",
               "교차로": " Intersection", "시장": " Market", "공원": " Park", "병원": " Hospital",
               "대학교": " Univ.", "대학": " Univ.", "초등학교": " Elementary School", "중학교": " Middle School",
               "고등학교": " High School", "구청": "-gu Office", "시청": " City Hall", "주민센터": " Community Center",
               "아파트": " Apt.", "마을": " Village", "해수욕장": " Beach", "공항": " Airport", "항": " Port",
               "다리": " Bridge", "대교": " Bridge", "IC": " IC", "JC": " JC"},
}


def load_roman_dict():
    try:
        d = json.loads(ROMAN_FILE.read_text(encoding="utf-8"))
        return {"exact": d.get("exact", {}), "suffix": d.get("suffix", {})}
    except Exception:
        return {"exact": dict(ROMAN_DEFAULT["exact"]), "suffix": dict(ROMAN_DEFAULT["suffix"])}


def save_default_roman_dict():
    if not ROMAN_FILE.exists():
        APP_DIR.mkdir(parents=True, exist_ok=True)
        ROMAN_FILE.write_text(json.dumps(ROMAN_DEFAULT, ensure_ascii=False, indent=2), encoding="utf-8")


def _cap(word):
    return word[:1].upper() + word[1:] if word else word


def romanize_name(name, dic=None):
    """정류장·지명 하나 → 영문. 사전 → 끝말 사전 → 로마자 표기 순."""
    import re
    dic = dic or load_roman_dict()
    name = name.strip()
    if not name:
        return ""
    if name in dic["exact"]:
        return dic["exact"][name]
    for suf in sorted(dic["suffix"], key=len, reverse=True):
        if name.endswith(suf) and len(name) > len(suf):
            stem = name[: -len(suf)]
            return romanize_name(stem, dic) + dic["suffix"][suf]
    parts = []
    for word in name.split():
        if word in dic["exact"]:
            parts.append(dic["exact"][word]); continue
        tokens = re.findall(r"[가-힣]+|[0-9]+|[^가-힣0-9]+", word)
        buf = ""
        for i, tk in enumerate(tokens):
            if re.match(r"[가-힣]", tk):
                if i > 0 and tokens[i - 1].isdigit() and tk in ("가", "로", "길", "동", "번길"):
                    buf += "-" + romanize_word(tk)          # 을지로4가 → Euljiro 4-ga
                else:
                    buf += (" " if buf and buf[-1].isdigit() else "") + _cap(romanize_word(tk))
            elif tk.isdigit():
                buf += (" " if buf and buf[-1].isalpha() else "") + tk
            else:
                buf += tk
        parts.append(buf)
    return " ".join(parts)


def romanize_via(via, dic=None):
    """경유지 칸: 한글 줄들을 영문 줄로 바꿔서 (한글 줄 + 영문 줄) 형태로 돌려줌"""
    dic = dic or load_roman_dict()
    ko, _ = parse_via(via)
    lines = [",".join(s for s in stops) for stops in ko]
    en = [",".join(romanize_name(s, dic) for s in stops) for stops in ko]
    return "\n".join(lines + en)


# ───────────────────────── BMP 저장 · 파일 이름 ─────────────────────────
BMP_DEPTHS = ["24비트 (기본)", "8비트 (256색)", "4비트 (16색)", "1비트 (흑백)"]


def save_bmp(img, path, depth="24비트 (기본)"):
    import struct
    img = img.convert("RGB")
    if depth.startswith("8"):
        img.quantize(colors=256, method=Image.Quantize.MEDIANCUT).save(path, "BMP")
    elif depth.startswith("1"):
        img.convert("L").point(lambda v: 255 if v > 40 else 0).convert("1").save(path, "BMP")
    elif depth.startswith("4"):
        q = img.quantize(colors=16, method=Image.Quantize.MEDIANCUT)
        pal = (q.getpalette() or [])[:48]
        pal += [0] * (48 - len(pal))
        w, h = q.size
        row_bytes = ((w + 1) // 2 + 3) & ~3
        px = q.load()
        data = bytearray()
        for y in range(h - 1, -1, -1):
            row = bytearray()
            for x in range(0, w, 2):
                a = px[x, y] & 15
                b = (px[x + 1, y] & 15) if x + 1 < w else 0
                row.append((a << 4) | b)
            row += b"\0" * (row_bytes - len(row))
            data += row
        palette = b"".join(bytes((pal[i * 3 + 2], pal[i * 3 + 1], pal[i * 3], 0)) for i in range(16))
        off = 14 + 40 + len(palette)
        hdr = struct.pack("<2sIHHI", b"BM", off + len(data), 0, 0, off)
        info = struct.pack("<IiiHHIIiiII", 40, w, h, 1, 4, 0, len(data), 2835, 2835, 16, 0)
        Path(path).write_bytes(hdr + info + palette + bytes(data))
    else:
        img.save(path, "BMP")


def safe_name(s):
    return "".join(ch for ch in str(s) if ch not in '\\/:*?"<>|\n\r\t').strip()


def file_names(route, n_pages, pattern="{no}_{no}"):
    """파일 이름 규칙. 쓸 수 있는 값: {no} {kind} {start_ko} {end_ko} {page} {template}"""
    pattern = (pattern or "{no}_{no}").strip()
    if pattern.lower().endswith(".bmp"):
        pattern = pattern[:-4]
    vals = _Missing({k: safe_name(v) for k, v in route.items() if isinstance(v, str)})
    vals["no"] = safe_name(route.get("no", "")) or "noname"
    names = []
    for i in range(n_pages):
        vals["page"] = str(i + 1)
        base = pattern.format_map(vals)
        if n_pages > 1 and "{page}" not in pattern:
            base += f"_{i + 1}"
        names.append((base or "noname") + ".bmp")
    return names


# ───────────────────────── 원본 BMP 비교 ─────────────────────────
def compare_images(orig, ours):
    """같은 크기 두 이미지 → (정확 일치율, 1칸 오차 허용 일치율, 겹쳐보기 이미지)"""
    import numpy as np
    a = np.asarray(orig.convert("RGB")).astype(int).max(2) > 60
    b = np.asarray(ours.convert("RGB")).astype(int).max(2) > 60
    if a.shape != b.shape:
        raise ValueError(f"크기가 다릅니다: 원본 {a.shape[1]}x{a.shape[0]}, 현재 {b.shape[1]}x{b.shape[0]}")
    def dil(m):
        o = m.copy()
        p = np.pad(m, 1)
        for dy in range(3):
            for dx in range(3):
                o |= p[dy:dy + m.shape[0], dx:dx + m.shape[1]]
        return o
    union = (a | b).sum()
    exact = (a & b).sum() / union if union else 1.0
    p = (b & dil(a)).sum() / max(1, b.sum())
    r = (a & dil(b)).sum() / max(1, a.sum())
    tol = 2 * p * r / (p + r) if p + r else 1.0
    ov = np.zeros(a.shape + (3,), np.uint8)
    ov[a & b] = (255, 220, 0)
    ov[a & ~b] = (255, 50, 50)
    ov[b & ~a] = (40, 210, 70)
    return exact, tol, Image.fromarray(ov, "RGB")


# ───────────────────────── 프로젝트 · 설정 ─────────────────────────
SETTINGS_FILE = APP_DIR / "settings.json"
DEFAULT_SETTINGS = {"bmp_depth": BMP_DEPTHS[0], "name_pattern": "{no}_{no}", "out_dir": str(OUT_DIR),
                    "usb_dir": "", "zoom": "맞춤", "kinds": {}}


def load_settings():
    s = dict(DEFAULT_SETTINGS)
    try:
        s.update(json.loads(SETTINGS_FILE.read_text(encoding="utf-8")))
    except Exception:
        pass
    return s


def save_settings(s):
    APP_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_FILE.write_text(json.dumps(s, ensure_ascii=False, indent=2), encoding="utf-8")


def save_project(path, routes, style, settings):
    data = {"app": "BusSignMaker", "version": APP_VERSION, "routes": routes, "style": style, "settings": settings}
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")


def load_project(path):
    d = json.loads(Path(path).read_text(encoding="utf-8"))
    return d.get("routes", []), d.get("style", {}), d.get("settings", {})


def generate_all(routes, style, settings, progress=None):
    """전체 생성 → (저장한 파일 수, 경고 목록 [(노선, 내용, 비율)])"""
    out_dir = Path(settings.get("out_dir") or OUT_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    n, warns = 0, []
    for i, r in enumerate(routes):
        pages = render_pages(r, style)
        for t, ratio in WARN:
            warns.append((r.get("no", ""), t, ratio))
        for img, name in zip(pages, file_names(r, len(pages), settings.get("name_pattern"))):
            save_bmp(img, out_dir / name, settings.get("bmp_depth", BMP_DEPTHS[0]))
            n += 1
        if progress:
            progress(i + 1, len(routes))
    return n, warns

# ───────────────────────── 엑셀 ─────────────────────────
XL_COLS = [("no", "번호"), ("start_ko", "기점(한)"), ("end_ko", "종점(한)"),
           ("start_en", "기점(영)"), ("end_en", "종점(영)"), ("via", "경유지"),
           ("template", "템플릿"), ("kind", "구분")]


def make_sample_excel(path):
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "노선"
    ws.append([c[1] for c in XL_COLS])
    ws.append(["360", "서울역 환승센터", "강일IC", "Seoul Station", "Gang-il IC",
               "동묘앞,을지로4가,\n강동역,십자성마을\nDongmyo,Euljuro-4ga,\nGangdong Stn,10JASUNG TOWN",
               "jinmyeong_front_192x160"])
    ws.append(["1100", "부산역", "하단역", "Busan Stn", "Hadan Stn",
               "남포동,자갈치\nNampo-dong,Jagalchi", "jinmyeong_front_192x160"])
    for c, w in zip("ABCDEFG", (8, 18, 14, 16, 14, 50, 26)):
        ws.column_dimensions[c].width = w
    wb.save(path)


def export_excel(path, routes):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "노선"
    ws.append([c[1] for c in XL_COLS])
    for r in routes:
        ws.append([r.get(k, "") for k, _ in XL_COLS])
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for c, w in zip("ABCDEFGH", (8, 18, 14, 16, 14, 50, 26, 8)):
        ws.column_dimensions[c].width = w
    wb.save(path)


def load_excel(path):
    from openpyxl import load_workbook
    ws = load_workbook(path, data_only=True).active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = [str(h or "").strip() for h in rows[0]]
    idx = {}
    for key, label in XL_COLS:
        if label in header:
            idx[key] = header.index(label)
    routes = []
    for row in rows[1:]:
        if not row or all(v in (None, "") for v in row):
            continue
        r = {k: ("" if row[i] is None else str(row[i])) for k, i in idx.items() if i < len(row)}
        if r.get("no", "").endswith(".0"):
            r["no"] = r["no"][:-2]
        r.setdefault("template", "jinmyeong_front_192x160")
        routes.append(r)
    return routes


def open_folder(path):
    path = str(path)
    if sys.platform == "darwin":
        subprocess.Popen(["open", path])
    elif os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    else:
        subprocess.Popen(["xdg-open", path])


# ───────────────────────── GUI ─────────────────────────
PEN_COLORS = ["흰색", "노랑", "빨강", "초록", "주황", "하늘", "끄기(검정)", "원래대로"]
GLYPH_SIZES_TEXT = {"14 · 두 줄 기점/종점": (14, 13, 7), "24 · 한 줄 한글 종점": (24, 23, 12),
                    "13 · 영문 한 줄 / 경유지": (13, 13, 7), "11 · 작은 글자": (11, 11, 6), "9 · 아주 작은 글자": (9, 9, 5)}
GLYPH_SIZES_NUM = {"22 · 큰 번호": 22, "19 · 작은 번호": 19, "24 · 후면 번호": 24}
IS_MAC = sys.platform == "darwin"
MOD = "Command" if IS_MAC else "Control"


def run_gui():
    import copy
    import shutil
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, simpledialog
    from PIL import ImageTk

    APP_DIR.mkdir(parents=True, exist_ok=True)
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    load_glyphs()
    save_default_roman_dict()
    write_example_templates()
    settings = load_settings()

    def load_presets():
        try:
            return json.loads(PRESET_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}

    root = tk.Tk()
    root.title(f"노선이미지 생성기 - EBASE BUS_Controller  ·  v{APP_VERSION}")
    root.geometry("1380x860")
    root.minsize(1000, 640)

    S = {  # 화면 상태
        "routes": [], "sel": None, "photo": None, "loading": False, "project": None, "dirty": False,
        "undo": [], "redo": [], "snap": None, "commit_job": None, "page": 0, "pages": [],
        "cmp_img": None, "cmp_name": "", "last_files": [], "paint_val": None,
    }
    status = tk.StringVar(value="준비")

    # ── 레이아웃 ──
    bar = ttk.Frame(root, padding=(4, 4))
    bar.pack(fill="x")
    ttk.Label(root, textvariable=status, anchor="w", relief="sunken").pack(fill="x", side="bottom")
    body = ttk.Frame(root)
    body.pack(fill="both", expand=True)

    # 노선 목록
    lf = ttk.LabelFrame(body, text="노선 목록", padding=4)
    lf.pack(side="left", fill="y", padx=4, pady=4)
    lbf = ttk.Frame(lf)
    lbf.pack(fill="both", expand=True)
    lb = tk.Listbox(lbf, width=18, exportselection=False)
    lb.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(lbf, command=lb.yview)
    sb.pack(side="left", fill="y")
    lb.config(yscrollcommand=sb.set)
    lbtn = ttk.Frame(lf)
    lbtn.pack(fill="x", pady=(4, 0))

    # 가운데 탭
    mid = ttk.Notebook(body, width=400)
    mid.pack(side="left", fill="y", padx=4, pady=4)
    tab_edit, tab_style, tab_font, tab_out = (ttk.Frame(mid, padding=8) for _ in range(4))
    mid.add(tab_edit, text="편집")
    mid.add(tab_style, text="스타일")
    mid.add(tab_font, text="폰트")
    mid.add(tab_out, text="출력")

    # ── 편집 탭 ──
    vars_ = {}
    fields = [("no", "번호"), ("start_ko", "기점(한)"), ("end_ko", "종점(한)"),
              ("start_en", "기점(영)"), ("end_en", "종점(영)")]
    for i, (k, label) in enumerate(fields):
        ttk.Label(tab_edit, text=label).grid(row=i, column=0, sticky="e", padx=3, pady=2)
        v = tk.StringVar()
        ttk.Entry(tab_edit, textvariable=v, width=30).grid(row=i, column=1, sticky="we", pady=2)
        vars_[k] = v
    ttk.Label(tab_edit, text="구분").grid(row=5, column=0, sticky="e", padx=3)
    kind_var = tk.StringVar(value="일반")
    ttk.Combobox(tab_edit, textvariable=kind_var, values=KINDS, state="readonly", width=10).grid(
        row=5, column=1, sticky="w", pady=2)
    ttk.Label(tab_edit, text="경유지").grid(row=6, column=0, sticky="ne", padx=3)
    via_txt = tk.Text(tab_edit, width=30, height=6, wrap="word", undo=False)
    via_txt.grid(row=6, column=1, sticky="we", pady=2)
    ttk.Label(tab_edit, text="쉼표 = 정류장, 줄바꿈 = 줄 (한글 줄 → 영문 줄)", foreground="gray").grid(
        row=7, column=1, sticky="w")
    ttk.Label(tab_edit, text="템플릿").grid(row=8, column=0, sticky="e", padx=3)
    tpl_var = tk.StringVar(value="jinmyeong_front_192x160")
    tpl_cb = ttk.Combobox(tab_edit, textvariable=tpl_var, values=all_template_names(), state="readonly", width=30)
    tpl_cb.grid(row=8, column=1, sticky="we", pady=2)
    rom_frame = ttk.LabelFrame(tab_edit, text="영문 자동 변환 (국어의 로마자 표기법 + 사전)", padding=6)
    rom_frame.grid(row=9, column=0, columnspan=2, sticky="we", pady=(10, 0))
    rom_over = tk.BooleanVar(value=False)
    ttk.Checkbutton(rom_frame, text="이미 적힌 영문도 덮어쓰기", variable=rom_over).pack(anchor="w")
    rf2 = ttk.Frame(rom_frame)
    rf2.pack(fill="x", pady=(4, 0))

    # ── 스타일 탭 ──
    style_vars = {}
    sf = ttk.LabelFrame(tab_style, text="색·모양", padding=6)
    sf.pack(fill="x")
    for i, (k, label, _) in enumerate(STYLE_KEYS):
        r_, c_ = divmod(i, 2)
        ttk.Label(sf, text=label).grid(row=r_, column=c_ * 2, sticky="e", padx=3, pady=2)
        values = (ARROW_SHAPES if k == "arrow_shape" else ALIGNS if k == "align"
                  else list(WEIGHTS) if k == "weight" else list(NUM_WEIGHTS) if k == "num_weight" else list(COLORS))
        v = tk.StringVar(value="(기본)")
        ttk.Combobox(sf, textvariable=v, values=values, state="readonly", width=10).grid(
            row=r_, column=c_ * 2 + 1, sticky="w", pady=2)
        style_vars[k] = v
    kf = ttk.LabelFrame(tab_style, text="노선 구분별 번호 표시 (급행·심야 등)", padding=6)
    kf.pack(fill="x", pady=(8, 0))
    kind_vars = {}
    for i, kd in enumerate(KINDS[1:]):
        c0, inv0 = KIND_DEFAULTS[kd]
        ttk.Label(kf, text=kd).grid(row=i, column=0, sticky="e", padx=3, pady=1)
        cv, iv = tk.StringVar(value=c0), tk.BooleanVar(value=inv0)
        ttk.Combobox(kf, textvariable=cv, values=list(COLORS), state="readonly", width=8).grid(row=i, column=1, sticky="w")
        ttk.Checkbutton(kf, text="반전(색 상자)", variable=iv).grid(row=i, column=2, sticky="w", padx=4)
        kind_vars[kd] = (cv, iv)
    pf_ = ttk.LabelFrame(tab_style, text="회사 프리셋 (스타일+폰트+구분 표시 저장)", padding=6)
    pf_.pack(fill="x", pady=(8, 0))
    preset_var = tk.StringVar()
    preset_cb = ttk.Combobox(pf_, textvariable=preset_var, values=list(load_presets()), width=16)
    preset_cb.pack(side="left")

    # ── 폰트 탭 ──
    font_map = {"(자동)": ""}
    font_map.update(scan_fonts())
    font_vars = {"text_font": tk.StringVar(value="(자동)"), "num_font": tk.StringVar(value="(자동)")}
    pixel_var = tk.BooleanVar(value=False)
    font_cbs = {}

    def path_to_name(path):
        if not path:
            return "(자동)"
        for n, pth in font_map.items():
            if pth == path:
                return n
        n = Path(path).name
        font_map[n] = path
        for cb in font_cbs.values():
            cb["values"] = list(font_map)
        return n

    def browse_font(key):
        p = filedialog.askopenfilename(title="폰트 파일 선택",
                                       filetypes=[("폰트", "*.ttf *.ttc *.otf"), ("모든 파일", "*.*")])
        if p:
            try:
                ImageFont.truetype(p, 16)
            except Exception as e:
                messagebox.showerror("폰트 오류", f"이 폰트는 열 수 없습니다.\n{e}")
                return
            font_vars[key].set(path_to_name(p))

    for i, (key, label) in enumerate([("text_font", "글자 폰트"), ("num_font", "번호 폰트")]):
        ttk.Label(tab_font, text=label).grid(row=i, column=0, sticky="e", padx=3, pady=3)
        cb = ttk.Combobox(tab_font, textvariable=font_vars[key], values=list(font_map), state="readonly", width=26)
        cb.grid(row=i, column=1, sticky="we", pady=3)
        font_cbs[key] = cb
        ttk.Button(tab_font, text="찾기…", width=6, command=lambda k=key: browse_font(k)).grid(row=i, column=2, padx=2)
    ttk.Checkbutton(tab_font, text="픽셀폰트 모드 (LED 비트맵 폰트용)", variable=pixel_var).grid(
        row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))
    ttk.Label(tab_font, text="(자동) = 앱에 들어 있는 나눔고딕 / 원본형 번호 글꼴\n"
                             "~/BusSignMaker/fonts 에 넣은 폰트는 목록에 자동으로 나옵니다.",
              foreground="gray", justify="left").grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

    # ── 출력 탭 ──
    out_vars = {"bmp_depth": tk.StringVar(value=settings["bmp_depth"]),
                "name_pattern": tk.StringVar(value=settings["name_pattern"]),
                "out_dir": tk.StringVar(value=settings["out_dir"]),
                "usb_dir": tk.StringVar(value=settings["usb_dir"])}
    ttk.Label(tab_out, text="BMP 형식").grid(row=0, column=0, sticky="e", padx=3, pady=3)
    ttk.Combobox(tab_out, textvariable=out_vars["bmp_depth"], values=BMP_DEPTHS, state="readonly", width=18).grid(
        row=0, column=1, sticky="w", pady=3)
    ttk.Label(tab_out, text="파일 이름").grid(row=1, column=0, sticky="e", padx=3, pady=3)
    ttk.Entry(tab_out, textvariable=out_vars["name_pattern"], width=26).grid(row=1, column=1, sticky="we", pady=3)
    ttk.Label(tab_out, text="쓸 수 있는 값: {no} {kind} {start_ko} {end_ko} {page}\n"
                            "예) {no}_{no} → 360_360.bmp, 여러 장이면 _1, _2가 붙음",
              foreground="gray", justify="left").grid(row=2, column=1, sticky="w")
    for i, (k, label) in enumerate([("out_dir", "출력 폴더"), ("usb_dir", "USB·복사 폴더")]):
        ttk.Label(tab_out, text=label).grid(row=3 + i, column=0, sticky="e", padx=3, pady=3)
        ttk.Entry(tab_out, textvariable=out_vars[k], width=26).grid(row=3 + i, column=1, sticky="we", pady=3)
        ttk.Button(tab_out, text="찾기…", width=6,
                   command=lambda k=k: (lambda d: d and out_vars[k].set(d))(filedialog.askdirectory())).grid(
            row=3 + i, column=2, padx=2)

    def save_out_settings(*_):
        for k, v in out_vars.items():
            settings[k] = v.get()
        try:
            save_settings(settings)
        except Exception:
            pass
        refresh_preview()

    for v in out_vars.values():
        v.trace_add("write", save_out_settings)

    # ── 미리보기 ──
    pv = ttk.LabelFrame(body, text="미리보기", padding=4)
    pv.pack(side="left", fill="both", expand=True, padx=4, pady=4)
    pctl = ttk.Frame(pv)
    pctl.pack(fill="x")
    zoom_var = tk.StringVar(value=str(settings.get("zoom", "맞춤")))
    ttk.Label(pctl, text="배율").pack(side="left")
    ttk.Combobox(pctl, textvariable=zoom_var, values=["맞춤", 2, 3, 4, 5, 6, 8], width=4, state="readonly").pack(side="left", padx=(2, 8))
    page_var = tk.StringVar(value="1")
    ttk.Label(pctl, text="페이지").pack(side="left")
    page_cb = ttk.Combobox(pctl, textvariable=page_var, values=["1"], width=3, state="readonly")
    page_cb.pack(side="left", padx=(2, 8))
    mode_var = tk.StringVar(value="보기")
    ttk.Radiobutton(pctl, text="보기", variable=mode_var, value="보기").pack(side="left")
    ttk.Radiobutton(pctl, text="점 편집", variable=mode_var, value="편집").pack(side="left")
    pen_var = tk.StringVar(value="흰색")
    ttk.Combobox(pctl, textvariable=pen_var, values=PEN_COLORS, width=9, state="readonly").pack(side="left", padx=(4, 4))
    cmp_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(pctl, text="원본 겹쳐보기", variable=cmp_var).pack(side="left", padx=(8, 0))
    cwrap = ttk.Frame(pv)
    cwrap.pack(fill="both", expand=True, pady=(4, 0))
    canvas = tk.Canvas(cwrap, background="#2b2b2b", highlightthickness=0)
    hsb = ttk.Scrollbar(cwrap, orient="horizontal", command=canvas.xview)
    vsb = ttk.Scrollbar(cwrap, orient="vertical", command=canvas.yview)
    canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)
    img_item = canvas.create_image(0, 0, anchor="nw")
    info = tk.StringVar()
    warn_var = tk.StringVar()
    ttk.Label(pv, textvariable=info, foreground="gray").pack(anchor="w", pady=(4, 0))
    ttk.Label(pv, textvariable=warn_var, foreground="#c0392b", wraplength=760, justify="left").pack(anchor="w")
    legend = ttk.Label(pv, text="", foreground="gray")
    legend.pack(anchor="w")

    def zoom_now(img_w=192, img_h=160):
        v = zoom_var.get()
        if v.isdigit():
            return int(v)
        cw = max(100, canvas.winfo_width() - 4)
        ch = max(100, canvas.winfo_height() - 4)
        return max(1, min(8, cw // max(1, img_w), ch // max(1, img_h)))

    # ── 데이터 ↔ 화면 ──
    def current_style():
        st = {k: v.get() for k, v in style_vars.items()}
        for k, v in font_vars.items():
            st[k] = font_map.get(v.get(), "")
        st["pixel_mode"] = "1" if pixel_var.get() else ""
        st["kinds"] = {kd: {"color": cv.get(), "invert": bool(iv.get())} for kd, (cv, iv) in kind_vars.items()}
        return st

    def apply_style(st):
        S["loading"] = True
        for k, v in style_vars.items():
            v.set(st.get(k, "(기본)") or "(기본)")
        for k, v in font_vars.items():
            v.set(path_to_name(st.get(k, "")))
        pixel_var.set(str(st.get("pixel_mode", "")) == "1")
        for kd, (cv, iv) in kind_vars.items():
            c0, i0 = KIND_DEFAULTS[kd]
            d = (st.get("kinds") or {}).get(kd, {})
            cv.set(d.get("color", c0))
            iv.set(bool(d.get("invert", i0)))
        S["loading"] = False

    def form_route():
        r = {k: v.get() for k, v in vars_.items()}
        r["via"] = via_txt.get("1.0", "end-1c")
        r["template"] = tpl_var.get()
        r["kind"] = kind_var.get()
        return r

    def cur():
        i = S["sel"]
        return S["routes"][i] if i is not None and i < len(S["routes"]) else None

    def label_of(r):
        k = r.get("kind")
        return (r.get("no") or "(새 노선)") + (f"  [{k}]" if k and k != "일반" else "")

    def refill_list(keep=None):
        lb.delete(0, "end")
        for r in S["routes"]:
            lb.insert("end", label_of(r))
        if S["routes"]:
            i = min(keep if keep is not None else (S["sel"] or 0), len(S["routes"]) - 1)
            S["sel"] = i
            lb.selection_clear(0, "end")
            lb.selection_set(i)
            lb.see(i)
        else:
            S["sel"] = None

    def load_form(i):
        S["loading"] = True
        S["sel"] = i
        r = S["routes"][i]
        for k, v in vars_.items():
            v.set(r.get(k, ""))
        via_txt.delete("1.0", "end")
        via_txt.insert("1.0", r.get("via", ""))
        via_txt.edit_modified(False)
        tpl_var.set(r.get("template") or "jinmyeong_front_192x160")
        kind_var.set(r.get("kind") or "일반")
        S["loading"] = False
        S["page"] = 0
        refresh_preview()

    # ── 되돌리기 ──
    def snapshot():
        return json.dumps({"routes": S["routes"], "sel": S["sel"]}, ensure_ascii=False)

    def commit(push=True):
        """현재 상태를 되돌리기 기록에 남김"""
        S["commit_job"] = None
        snap = snapshot()
        if push and S["snap"] is not None and snap != S["snap"]:
            S["undo"].append(S["snap"])
            del S["undo"][:-200]
            S["redo"].clear()
            S["dirty"] = True
            update_title()
        S["snap"] = snap

    def schedule_commit():
        if S["commit_job"]:
            root.after_cancel(S["commit_job"])
        S["commit_job"] = root.after(700, commit)

    def restore(snap):
        d = json.loads(snap)
        S["routes"] = d["routes"]
        refill_list(d.get("sel") or 0)
        if S["routes"]:
            load_form(S["sel"])
        else:
            refresh_preview()
        S["snap"] = snap
        S["dirty"] = True
        update_title()

    def undo(*_):
        if S["commit_job"]:
            root.after_cancel(S["commit_job"]); commit()
        if not S["undo"]:
            status.set("되돌릴 작업이 없습니다"); return
        S["redo"].append(snapshot())
        restore(S["undo"].pop())
        status.set("되돌렸습니다")

    def redo(*_):
        if not S["redo"]:
            status.set("다시 실행할 작업이 없습니다"); return
        S["undo"].append(snapshot())
        restore(S["redo"].pop())
        status.set("다시 실행했습니다")

    def update_title():
        name = Path(S["project"]).name if S["project"] else "새 프로젝트"
        root.title(f"{'● ' if S['dirty'] else ''}{name} — 노선이미지 생성기 v{APP_VERSION}")

    # ── 미리보기 그리기 ──
    def render_current():
        r = cur()
        if r is None:
            return []
        pages = render_pages(r, current_style())
        S["warns"] = list(WARN)
        return pages

    def refresh_preview(*_):
        if S["loading"]:
            return
        r = cur()
        if r is None:
            canvas.itemconfig(img_item, image="")
            info.set("노선이 없습니다. [새 노선] 또는 엑셀/프로젝트를 불러오세요.")
            return
        try:
            pages = render_current()
        except Exception as e:  # noqa
            status.set(f"렌더 오류: {e}")
            return
        S["pages"] = pages
        n = len(pages)
        page_cb["values"] = [str(i + 1) for i in range(n)]
        S["page"] = min(S["page"], n - 1)
        if page_var.get() != str(S["page"] + 1):
            S["loading"] = True; page_var.set(str(S["page"] + 1)); S["loading"] = False
        img = pages[S["page"]]
        z = zoom_now(img.width, img.height)
        S["z"] = z
        extra = ""
        if cmp_var.get() and S["cmp_img"] is not None:
            try:
                ex, tol, ov = compare_images(S["cmp_img"], img)
                img = ov
                extra = f"  |  원본 비교({S['cmp_name']}): 정확 {ex * 100:.1f}% · 1칸 허용 {tol * 100:.1f}%"
                legend.config(text="겹쳐보기: 노랑 = 같음, 빨강 = 원본에만 있음, 초록 = 지금에만 있음")
            except Exception as e:
                extra = f"  |  비교 불가: {e}"
        else:
            legend.config(text="점 편집: 왼쪽 클릭·드래그 = 칠하기, 오른쪽 클릭 = 원래대로" if mode_var.get() == "편집" else "")
        big = img.resize((img.width * z, img.height * z), Image.NEAREST)
        if mode_var.get() == "편집" and z >= 3:     # 편집할 때는 격자를 얇게 표시
            d = ImageDraw.Draw(big)
            for x in range(0, big.width, z):
                d.line([(x, 0), (x, big.height)], fill=(40, 40, 40))
            for y in range(0, big.height, z):
                d.line([(0, y), (big.width, y)], fill=(40, 40, 40))
        S["photo"] = ImageTk.PhotoImage(big)
        canvas.itemconfig(img_item, image=S["photo"])
        canvas.configure(scrollregion=(0, 0, big.width, big.height))
        names = file_names(r, n, settings.get("name_pattern"))
        info.set(f"{img.width}x{img.height} px  |  {r.get('template')}  |  {names[S['page']]}"
                 f"{'  |  편집한 점 ' + str(len((r.get('edits') or {}).get(str(S['page']), {}))) + '개' if (r.get('edits') or {}).get(str(S['page'])) else ''}{extra}")
        ws = [w for w in S.get("warns", []) if w[1] > 0]
        warn_var.set("⚠ 압축됨: " + " / ".join(f"{t} ({ratio * 100:.0f}%)" for t, ratio in ws) if ws else "")

    def on_form_change(*_):
        if S["loading"]:
            return
        r = cur()
        if r is None:
            return
        r.update(form_route())
        i = S["sel"]
        lb.delete(i)
        lb.insert(i, label_of(r))
        lb.selection_set(i)
        refresh_preview()
        schedule_commit()

    def on_style_change(*_):
        if not S["loading"]:
            refresh_preview()

    def on_page(*_):
        if S["loading"]:
            return
        try:
            S["page"] = int(page_var.get()) - 1
        except ValueError:
            S["page"] = 0
        refresh_preview()

    # ── 점 편집 ──
    def canvas_xy(ev):
        z = S.get("z", 4)
        x = int(canvas.canvasx(ev.x) // z)
        y = int(canvas.canvasy(ev.y) // z)
        return x, y

    def paint(ev, value):
        r = cur()
        if r is None or mode_var.get() != "편집" or not S["pages"]:
            return
        x, y = canvas_xy(ev)
        img = S["pages"][S["page"]]
        if not (0 <= x < img.width and 0 <= y < img.height):
            return
        ed = r.setdefault("edits", {}).setdefault(str(S["page"]), {})
        key = f"{x},{y}"
        if value == "원래대로":
            if key not in ed:
                return
            ed.pop(key)
        else:
            rgb = [0, 0, 0] if value.startswith("끄기") else list(COLORS.get(value) or (255, 255, 255))
            if ed.get(key) == rgb:
                return
            ed[key] = rgb
        if not ed:
            r["edits"].pop(str(S["page"]), None)
        refresh_preview()

    def on_press(ev):
        if mode_var.get() != "편집":
            return
        S["paint_val"] = pen_var.get()
        paint(ev, S["paint_val"])

    def on_drag(ev):
        if S["paint_val"]:
            paint(ev, S["paint_val"])

    def on_release(_):
        if S["paint_val"]:
            S["paint_val"] = None
            commit()

    def on_right(ev):
        if mode_var.get() == "편집":
            paint(ev, "원래대로")
            commit()

    canvas.bind("<ButtonPress-1>", on_press)
    canvas.bind("<B1-Motion>", on_drag)
    canvas.bind("<ButtonRelease-1>", on_release)
    canvas.bind("<Button-2>" if IS_MAC else "<Button-3>", on_right)
    canvas.bind("<Control-Button-1>", on_right)

    def clear_edits():
        r = cur()
        if r and (r.get("edits") or {}).get(str(S["page"])):
            if messagebox.askyesno("점 편집 지우기", "이 페이지에서 손으로 고친 점을 모두 지울까요?"):
                r["edits"].pop(str(S["page"]), None)
                refresh_preview(); commit()

    # ── 노선 목록 버튼 ──
    def on_list(_):
        s = lb.curselection()
        if s and s[0] != S["sel"]:
            if S["commit_job"]:
                root.after_cancel(S["commit_job"]); commit()
            load_form(s[0])

    def new_route():
        S["routes"].append({"no": "", "template": tpl_var.get() or "jinmyeong_front_192x160", "kind": "일반"})
        refill_list(len(S["routes"]) - 1)
        load_form(S["sel"]); commit()

    def dup_route():
        r = cur()
        if r is None:
            return
        S["routes"].insert(S["sel"] + 1, copy.deepcopy(r))
        refill_list(S["sel"] + 1)
        load_form(S["sel"]); commit()
        status.set("노선을 복제했습니다")

    def del_route():
        r = cur()
        if r is None:
            return
        if not messagebox.askyesno("노선 삭제", f"'{label_of(r)}' 노선을 삭제할까요? (되돌리기 가능)"):
            return
        S["routes"].pop(S["sel"])
        refill_list(max(0, S["sel"] - 1))
        if S["routes"]:
            load_form(S["sel"])
        else:
            refresh_preview()
        commit()

    def move_route(d):
        i = S["sel"]
        if i is None or not (0 <= i + d < len(S["routes"])):
            return
        S["routes"][i], S["routes"][i + d] = S["routes"][i + d], S["routes"][i]
        refill_list(i + d)
        commit()

    for t, c in [("새 노선", new_route), ("복제", dup_route), ("삭제", del_route),
                 ("▲", lambda: move_route(-1)), ("▼", lambda: move_route(1))]:
        ttk.Button(lbtn, text=t, width=5 if len(t) > 1 else 2, command=c).pack(side="left", padx=1)

    # ── 영문 자동 변환 ──
    def fill_english(r, overwrite):
        dic = load_roman_dict()
        changed = False
        for ko, en in (("start_ko", "start_en"), ("end_ko", "end_en")):
            if r.get(ko) and (overwrite or not r.get(en, "").strip()):
                r[en] = romanize_name(r[ko], dic); changed = True
        kos, ens = parse_via(r.get("via", ""))
        if kos and (overwrite or not ens):
            r["via"] = romanize_via(r.get("via", ""), dic); changed = True
        return changed

    def romanize_current():
        r = cur()
        if r is None:
            return
        r.update(form_route())
        if fill_english(r, rom_over.get()):
            load_form(S["sel"]); commit()
            status.set("영문을 자동으로 채웠습니다. 표기가 맞는지 확인해 주세요.")
        else:
            status.set("채울 빈 영문 칸이 없습니다 ('덮어쓰기'를 켜면 다시 만듭니다)")

    def romanize_all():
        n = sum(fill_english(r, rom_over.get()) for r in S["routes"])
        if S["routes"]:
            load_form(S["sel"])
        commit()
        status.set(f"{n}개 노선의 영문을 채웠습니다")

    def open_roman_dict():
        save_default_roman_dict()
        open_folder(ROMAN_FILE)
        status.set(f"사전 파일: {ROMAN_FILE} (고친 뒤 저장하면 바로 적용)")

    ttk.Button(rf2, text="이 노선 채우기", command=romanize_current).pack(side="left")
    ttk.Button(rf2, text="모든 노선 채우기", command=romanize_all).pack(side="left", padx=4)
    ttk.Button(rf2, text="사전 편집…", command=open_roman_dict).pack(side="left")

    # ── 프리셋 ──
    def apply_preset(*_):
        p = load_presets().get(preset_var.get())
        if p:
            apply_style(p)
            refresh_preview()

    def save_preset():
        name = preset_var.get().strip()
        if not name:
            messagebox.showinfo("프리셋", "왼쪽 칸에 프리셋(회사) 이름을 입력하세요.")
            return
        data = load_presets()
        data[name] = current_style()
        PRESET_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        preset_cb["values"] = list(data)
        status.set(f"프리셋 저장: {name}")

    preset_cb.bind("<<ComboboxSelected>>", apply_preset)
    ttk.Button(pf_, text="프리셋 저장", command=save_preset).pack(side="left", padx=4)

    # ── 파일: 프로젝트·엑셀 ──
    def ask_discard():
        if not S["dirty"]:
            return True
        a = messagebox.askyesnocancel("저장 확인", "저장하지 않은 변경이 있습니다. 저장할까요?")
        if a is None:
            return False
        if a:
            return save_project_cmd()
        return True

    def set_routes(routes, msg):
        S["routes"] = routes
        S["undo"].clear(); S["redo"].clear()
        refill_list(0)
        if routes:
            load_form(0)
        else:
            refresh_preview()
        S["snap"] = snapshot()
        status.set(msg)

    def new_project():
        if not ask_discard():
            return
        S["project"] = None
        set_routes([], "새 프로젝트")
        S["dirty"] = False; update_title()

    def open_project():
        if not ask_discard():
            return
        p = filedialog.askopenfilename(filetypes=[("노선 프로젝트", "*.bsm"), ("모든 파일", "*.*")])
        if not p:
            return
        try:
            routes, st, se = load_project(p)
        except Exception as e:
            messagebox.showerror("열기 오류", str(e)); return
        if st:
            apply_style(st)
        for k in ("bmp_depth", "name_pattern"):
            if se.get(k):
                out_vars[k].set(se[k])
        S["project"] = p
        set_routes(routes, f"프로젝트 열기: {p}")
        S["dirty"] = False; update_title()

    def save_project_cmd(as_new=False):
        if S["commit_job"]:
            root.after_cancel(S["commit_job"]); commit()
        p = S["project"]
        if as_new or not p:
            p = filedialog.asksaveasfilename(defaultextension=".bsm", initialfile="노선프로젝트.bsm",
                                             filetypes=[("노선 프로젝트", "*.bsm")])
            if not p:
                return False
        try:
            save_project(p, S["routes"], current_style(), {k: v.get() for k, v in out_vars.items()})
        except Exception as e:
            messagebox.showerror("저장 오류", str(e)); return False
        S["project"] = p
        S["dirty"] = False; update_title()
        status.set(f"프로젝트 저장: {p}")
        return True

    def load_xl():
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not p:
            return
        try:
            rs = load_excel(p)
        except Exception as e:
            messagebox.showerror("엑셀 오류", str(e)); return
        if S["routes"] and not messagebox.askyesno("엑셀 불러오기", f"{len(rs)}개 노선을 불러옵니다.\n지금 목록을 바꿀까요? (아니오 = 뒤에 추가)"):
            set_routes(S["routes"] + rs, f"{len(rs)}개 노선 추가")
        else:
            set_routes(rs, f"{len(rs)}개 노선 불러옴")
        S["dirty"] = True; update_title()

    def export_xl():
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="노선목록.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            export_excel(p, S["routes"])
            status.set(f"엑셀로 내보냄: {p} (점 편집 내용은 프로젝트 파일에만 저장됩니다)")

    def sample_xl():
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="노선샘플.xlsx",
                                         filetypes=[("Excel", "*.xlsx")])
        if p:
            make_sample_excel(p)
            status.set(f"샘플 저장: {p}")

    # ── 출력 ──
    def out_dir():
        d = Path(out_vars["out_dir"].get() or OUT_DIR)
        d.mkdir(parents=True, exist_ok=True)
        return d

    def save_current():
        r = cur()
        if r is None:
            return
        pages = render_pages(r, current_style())
        names = file_names(r, len(pages), out_vars["name_pattern"].get())
        for img, n in zip(pages, names):
            save_bmp(img, out_dir() / n, out_vars["bmp_depth"].get())
        S["last_files"] = [out_dir() / n for n in names]
        status.set(f"저장: {', '.join(names)}  →  {out_dir()}")

    def save_all():
        if not S["routes"]:
            messagebox.showinfo("전체 생성", "노선이 없습니다."); return
        st = current_style()
        files, warns = [], []
        for i, r in enumerate(S["routes"]):
            pages = render_pages(r, st)
            warns += [(r.get("no", ""), t, ratio) for t, ratio in WARN if ratio > 0]
            for img, n in zip(pages, file_names(r, len(pages), out_vars["name_pattern"].get())):
                save_bmp(img, out_dir() / n, out_vars["bmp_depth"].get())
                files.append(out_dir() / n)
            status.set(f"생성 중… {i + 1}/{len(S['routes'])}")
            root.update_idletasks()
        S["last_files"] = files
        status.set(f"{len(files)}개 파일 생성 완료 → {out_dir()}")
        refresh_preview()
        if warns:
            show_warnings(warns, len(files))

    def show_warnings(warns, nfiles):
        w = tk.Toplevel(root)
        w.title("압축 경고")
        w.geometry("640x380")
        ttk.Label(w, text=f"{nfiles}개 파일을 만들었습니다. 아래 {len(warns)}건은 글자가 칸보다 길어 좁게 눌렸습니다.\n"
                          "숫자가 작을수록 많이 눌린 것입니다. 노선을 더블클릭하면 그 노선으로 이동합니다.",
                  padding=8).pack(anchor="w")
        tv = ttk.Treeview(w, columns=("no", "text", "ratio"), show="headings")
        for c, t, wd in (("no", "노선", 80), ("text", "내용", 420), ("ratio", "남은 폭", 80)):
            tv.heading(c, text=t)
            tv.column(c, width=wd, anchor="w" if c == "text" else "center")
        for no, t, ratio in sorted(warns, key=lambda x: x[2]):
            tv.insert("", "end", values=(no, t, f"{ratio * 100:.0f}%"))
        tv.pack(fill="both", expand=True, padx=8)

        def jump(_):
            it = tv.focus()
            if not it:
                return
            no = tv.item(it)["values"][0]
            for i, r in enumerate(S["routes"]):
                if str(r.get("no")) == str(no):
                    refill_list(i); load_form(i); break

        tv.bind("<Double-1>", jump)

        def save_txt():
            p = filedialog.asksaveasfilename(defaultextension=".txt", initialfile="압축경고.txt")
            if p:
                Path(p).write_text("\n".join(f"{no}\t{t}\t{ratio * 100:.0f}%" for no, t, ratio in warns), encoding="utf-8")

        bf = ttk.Frame(w, padding=8)
        bf.pack(fill="x")
        ttk.Button(bf, text="목록 저장…", command=save_txt).pack(side="left")
        ttk.Button(bf, text="닫기", command=w.destroy).pack(side="right")

    def copy_usb():
        dest = out_vars["usb_dir"].get().strip()
        if not dest or not Path(dest).exists():
            dest = filedialog.askdirectory(title="복사할 USB(또는 폴더)를 고르세요")
            if not dest:
                return
            out_vars["usb_dir"].set(dest)
        files = [f for f in S["last_files"] if Path(f).exists()] or sorted(out_dir().glob("*.bmp"))
        if not files:
            messagebox.showinfo("복사", "복사할 BMP가 없습니다. 먼저 [전체 생성]을 하세요."); return
        if not messagebox.askyesno("USB로 복사", f"{len(files)}개 BMP를\n{dest}\n로 복사할까요? (같은 이름은 덮어씀)"):
            return
        try:
            for f in files:
                shutil.copy2(f, Path(dest) / Path(f).name)
        except Exception as e:
            messagebox.showerror("복사 오류", str(e)); return
        status.set(f"{len(files)}개 파일을 {dest} 로 복사했습니다")

    # ── 원본 비교 ──
    def load_compare():
        p = filedialog.askopenfilename(title="비교할 원본 BMP", filetypes=[("이미지", "*.bmp *.png"), ("모든 파일", "*.*")])
        if not p:
            return
        try:
            im = Image.open(p).convert("RGB")
        except Exception as e:
            messagebox.showerror("비교", str(e)); return
        S["cmp_img"], S["cmp_name"] = im, Path(p).name
        cmp_var.set(True)
        refresh_preview()
        status.set(f"원본 비교: {p}")

    # ── 글자 보정 창 ──
    def glyph_editor():
        w = tk.Toplevel(root)
        w.title("글자 보정 — 한 번 고치면 그 글자가 나오는 모든 노선에 적용")
        w.geometry("900x620")
        left = ttk.Frame(w, padding=6)
        left.pack(side="left", fill="y")
        ttk.Label(left, text="저장된 보정").pack(anchor="w")
        glb = tk.Listbox(left, width=26, exportselection=False)
        glb.pack(fill="y", expand=True)
        right = ttk.Frame(w, padding=6)
        right.pack(side="left", fill="both", expand=True)
        top = ttk.Frame(right)
        top.pack(fill="x")
        kind_v = tk.StringVar(value="TEXT")
        ttk.Radiobutton(top, text="글자", variable=kind_v, value="TEXT").pack(side="left")
        ttk.Radiobutton(top, text="번호", variable=kind_v, value="NUM").pack(side="left", padx=(0, 8))
        size_v = tk.StringVar(value=list(GLYPH_SIZES_TEXT)[0])
        size_cb = ttk.Combobox(top, textvariable=size_v, values=list(GLYPH_SIZES_TEXT), state="readonly", width=24)
        size_cb.pack(side="left")
        ttk.Label(top, text="  글자").pack(side="left")
        ch_v = tk.StringVar(value="서")
        ttk.Entry(top, textvariable=ch_v, width=4).pack(side="left")
        ttk.Button(top, text="불러오기", command=lambda: load_glyph()).pack(side="left", padx=4)
        ttk.Label(top, text="  폭").pack(side="left")
        wid_v = tk.IntVar(value=12)
        ttk.Spinbox(top, from_=1, to=60, textvariable=wid_v, width=4, command=lambda: resize_grid()).pack(side="left")
        gc = tk.Canvas(right, background="#111", highlightthickness=0)
        gc.pack(fill="both", expand=True, pady=6)
        msg = tk.StringVar(value="칸을 클릭·드래그해서 점을 켜고 끄세요. 저장하면 미리보기에 바로 반영됩니다.")
        ttk.Label(right, textvariable=msg, foreground="gray").pack(anchor="w")
        bot = ttk.Frame(right)
        bot.pack(fill="x", pady=(6, 0))
        G = {"m": None, "h": 14, "paint": None, "cell": 20}

        def keyinfo():
            k = kind_v.get()
            if k == "NUM":
                h = GLYPH_SIZES_NUM.get(size_v.get(), 19)
                return k, h, None
            h, full, half = GLYPH_SIZES_TEXT.get(size_v.get(), (14, 13, 7))
            return k, h, (full if is_wide(ch_v.get()[:1] or "가") else half) - 1

        def on_kind(*_):
            vals = list(GLYPH_SIZES_NUM) if kind_v.get() == "NUM" else list(GLYPH_SIZES_TEXT)
            size_cb["values"] = vals
            if size_v.get() not in vals:
                size_v.set(vals[0])

        kind_v.trace_add("write", on_kind)

        def draw():
            import numpy as np
            m = G["m"]
            gc.delete("all")
            if m is None:
                return
            h, wd = m.shape
            c = max(8, min(28, int(min(560 / max(1, h), 600 / max(1, wd)))))
            G["cell"] = c
            for y in range(h):
                for x in range(wd):
                    gc.create_rectangle(x * c, y * c, x * c + c - 1, y * c + c - 1,
                                        fill="#ffd400" if m[y, x] else "#262626", outline="#333")
            # 실제 크기 미리보기
            prev = bool_to_rgba(m, (255, 212, 0)).resize((wd * 3, h * 3), Image.NEAREST)
            G["photo"] = ImageTk.PhotoImage(prev)
            gc.create_image(wd * c + 20, 0, anchor="nw", image=G["photo"])
            gc.create_text(wd * c + 20, h * 3 + 8, anchor="nw", fill="#aaa", text="실제 크기 ×3")

        def load_glyph():
            import numpy as np
            ch = ch_v.get()[:1]
            if not ch:
                return
            render_route({"no": ""}, current_style())
            k, h, room = keyinfo()
            ov = glyph_override(k, h, ch)
            if ov is not None:
                m = ov.copy(); msg.set(f"저장된 보정을 불러왔습니다 ({glyph_key(k, h, ch)})")
            else:
                m = default_glyph(k, h, ch, room or 40)
                if m.shape[0] != h:
                    mm = np.zeros((h, m.shape[1]), bool); mm[:min(h, m.shape[0])] = m[:h]; m = mm
                msg.set("기본 모양을 불러왔습니다. 고친 뒤 [저장]을 누르세요.")
            G["m"], G["h"] = m, h
            wid_v.set(m.shape[1])
            draw()

        def resize_grid():
            import numpy as np
            m = G["m"]
            if m is None:
                return
            wd = max(1, int(wid_v.get() or 1))
            n = np.zeros((m.shape[0], wd), bool)
            n[:, :min(wd, m.shape[1])] = m[:, :wd]
            G["m"] = n
            draw()

        def cell_at(ev):
            c = G["cell"]
            return ev.y // c, ev.x // c

        def g_press(ev):
            m = G["m"]
            if m is None:
                return
            y, x = cell_at(ev)
            if 0 <= y < m.shape[0] and 0 <= x < m.shape[1]:
                G["paint"] = not m[y, x]
                m[y, x] = G["paint"]
                draw()

        def g_drag(ev):
            m = G["m"]
            if m is None or G["paint"] is None:
                return
            y, x = cell_at(ev)
            if 0 <= y < m.shape[0] and 0 <= x < m.shape[1] and m[y, x] != G["paint"]:
                m[y, x] = G["paint"]
                draw()

        gc.bind("<ButtonPress-1>", g_press)
        gc.bind("<B1-Motion>", g_drag)
        gc.bind("<ButtonRelease-1>", lambda e: G.update(paint=None))

        def refresh_list():
            glb.delete(0, "end")
            for k in sorted(GLYPHS):
                glb.insert("end", k)

        def pick(_):
            s = glb.curselection()
            if not s:
                return
            key = glb.get(s[0])
            font, h, ch = key.split("|", 2)
            kind_v.set("NUM" if font == "NUM" else "TEXT")
            sizes = GLYPH_SIZES_NUM if font == "NUM" else GLYPH_SIZES_TEXT
            for lab, v in sizes.items():
                if (v if font == "NUM" else v[0]) == int(h):
                    size_v.set(lab)
            ch_v.set(ch)
            load_glyph()

        glb.bind("<<ListboxSelect>>", pick)

        def save_glyph():
            m = G["m"]
            ch = ch_v.get()[:1]
            if m is None or not ch:
                return
            render_route({"no": ""}, current_style())
            k, h, _ = keyinfo()
            GLYPHS[glyph_key(k, h, ch)] = ["".join("#" if v else "." for v in row) for row in m]
            save_glyphs()
            refresh_list()
            refresh_preview()
            msg.set(f"저장했습니다: {glyph_key(k, h, ch)}")

        def reset_glyph():
            ch = ch_v.get()[:1]
            render_route({"no": ""}, current_style())
            k, h, _ = keyinfo()
            if GLYPHS.pop(glyph_key(k, h, ch), None) is not None:
                save_glyphs()
                refresh_list()
                refresh_preview()
            load_glyph()
            msg.set("보정을 지우고 기본 모양으로 되돌렸습니다.")

        ttk.Button(bot, text="저장", command=save_glyph).pack(side="left")
        ttk.Button(bot, text="기본값으로 되돌리기", command=reset_glyph).pack(side="left", padx=6)
        ttk.Button(bot, text="모두 지우기", command=lambda: (G["m"].fill(False), draw()) if G["m"] is not None else None).pack(side="left")
        ttk.Button(bot, text="닫기", command=w.destroy).pack(side="right")
        refresh_list()
        load_glyph()

    def reload_templates():
        tpl_cb["values"] = all_template_names()
        refresh_preview()
        status.set(f"템플릿 {len(all_template_names())}개 (폴더: {TEMPLATE_DIR})")

    # ── 메뉴 · 도구 모음 ──
    menubar = tk.Menu(root)
    m_file = tk.Menu(menubar, tearoff=0)
    acc = "⌘" if IS_MAC else "Ctrl+"
    m_file.add_command(label="새 프로젝트", command=new_project)
    m_file.add_command(label="프로젝트 열기…", accelerator=f"{acc}O", command=open_project)
    m_file.add_command(label="프로젝트 저장", accelerator=f"{acc}S", command=save_project_cmd)
    m_file.add_command(label="다른 이름으로 저장…", command=lambda: save_project_cmd(True))
    m_file.add_separator()
    m_file.add_command(label="엑셀 불러오기…", command=load_xl)
    m_file.add_command(label="엑셀로 내보내기…", command=export_xl)
    m_file.add_command(label="샘플 엑셀 만들기…", command=sample_xl)
    menubar.add_cascade(label="파일", menu=m_file)
    m_edit = tk.Menu(menubar, tearoff=0)
    m_edit.add_command(label="되돌리기", accelerator=f"{acc}Z", command=undo)
    m_edit.add_command(label="다시 실행", accelerator=f"{'⇧⌘' if IS_MAC else 'Ctrl+Y / Ctrl+Shift+'}Z", command=redo)
    m_edit.add_separator()
    m_edit.add_command(label="새 노선", command=new_route)
    m_edit.add_command(label="노선 복제", accelerator=f"{acc}D", command=dup_route)
    m_edit.add_command(label="노선 삭제", command=del_route)
    m_edit.add_separator()
    m_edit.add_command(label="이 페이지 점 편집 모두 지우기", command=clear_edits)
    menubar.add_cascade(label="편집", menu=m_edit)
    m_tool = tk.Menu(menubar, tearoff=0)
    m_tool.add_command(label="글자 보정…", command=glyph_editor)
    m_tool.add_command(label="원본 BMP 비교…", command=load_compare)
    m_tool.add_separator()
    m_tool.add_command(label="영문 채우기 (이 노선)", command=romanize_current)
    m_tool.add_command(label="영문 채우기 (모든 노선)", command=romanize_all)
    m_tool.add_command(label="영문 사전 편집…", command=open_roman_dict)
    m_tool.add_separator()
    m_tool.add_command(label="템플릿 폴더 열기", command=lambda: open_folder(TEMPLATE_DIR))
    m_tool.add_command(label="템플릿 다시 읽기", command=reload_templates)
    menubar.add_cascade(label="도구", menu=m_tool)
    m_out = tk.Menu(menubar, tearoff=0)
    m_out.add_command(label="현재 저장(BMP)", command=save_current)
    m_out.add_command(label="전체 생성", command=save_all)
    m_out.add_command(label="출력 폴더 열기", command=lambda: open_folder(out_dir()))
    m_out.add_command(label="USB·폴더로 복사…", command=copy_usb)
    menubar.add_cascade(label="출력", menu=m_out)
    root.config(menu=menubar)

    for text, cmd in [("열기", open_project), ("저장", save_project_cmd), ("엑셀 불러오기", load_xl),
                      ("현재 저장(BMP)", save_current), ("전체 생성", save_all),
                      ("출력폴더 열기", lambda: open_folder(out_dir())), ("USB로 복사", copy_usb),
                      ("원본 비교", load_compare), ("글자 보정", glyph_editor), ("↶", undo), ("↷", redo)]:
        ttk.Button(bar, text=text, command=cmd, width=3 if len(text) == 1 else None).pack(side="left", padx=2)

    for seq, fn in [("o", open_project), ("s", save_project_cmd), ("z", undo), ("d", dup_route)]:
        root.bind_all(f"<{MOD}-{seq}>", lambda e, f=fn: (f(), "break")[1])
    root.bind_all(f"<{MOD}-Shift-Z>", lambda e: (redo(), "break")[1])
    root.bind_all(f"<{MOD}-y>", lambda e: (redo(), "break")[1])

    # ── 변경 감지 ──
    lb.bind("<<ListboxSelect>>", on_list)
    for v in list(vars_.values()) + [tpl_var, kind_var]:
        v.trace_add("write", on_form_change)
    via_txt.bind("<<Modified>>", lambda e: (via_txt.edit_modified() and (via_txt.edit_modified(False), on_form_change())))
    for v in (list(style_vars.values()) + list(font_vars.values()) + [pixel_var]
              + [x for pair in kind_vars.values() for x in pair]):
        v.trace_add("write", on_style_change)
    zoom_var.trace_add("write", lambda *_: (settings.update(zoom=zoom_var.get()), refresh_preview()))
    S["last_size"] = None

    def on_canvas_resize(ev):
        if zoom_var.get() == "맞춤" and S["last_size"] != (ev.width, ev.height):
            S["last_size"] = (ev.width, ev.height)
            root.after_idle(refresh_preview)

    canvas.bind("<Configure>", on_canvas_resize)
    page_var.trace_add("write", on_page)
    mode_var.trace_add("write", refresh_preview)
    cmp_var.trace_add("write", lambda *_: (load_compare() if cmp_var.get() and S["cmp_img"] is None else refresh_preview()))

    def on_close():
        if ask_discard():
            try:
                save_settings(settings)
            except Exception:
                pass
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    # 시작 시 예시 노선
    S["routes"] = [{"no": "360", "start_ko": "서울역 환승센터", "end_ko": "강일IC",
                    "start_en": "Seoul Station", "end_en": "Gang-il IC",
                    "via": "동묘앞,을지로4가,\n강동역,십자성마을\nDongmyo,Euljuro-4ga,\nGangdong Stn,10JASUNG TOWN",
                    "template": "jinmyeong_front_192x160", "kind": "일반"}]
    refill_list(0)
    load_form(0)
    S["snap"] = snapshot()
    update_title()
    if not TEXT_FONT:
        messagebox.showwarning("폰트", "한글 폰트를 찾지 못했습니다.")
    if os.environ.get("BSM_TEST"):          # 자동 시험용 (일반 실행에는 영향 없음)
        globals()["_GUI"] = dict(locals())
    root.mainloop()


# ───────────────────────── 자체 검사 ─────────────────────────
SELFTEST_ROUTES = [
    {"no": "360", "start_ko": "서울역 환승센터", "end_ko": "강일IC", "start_en": "Seoul Station", "end_en": "Gang-il IC",
     "via": "동묘앞,을지로4가,\n강동역,십자성마을\nDongmyo,Euljuro-4ga,\nGangdong Stn,10JASUNG TOWN"},
    {"no": "1100", "start_ko": "부산역", "end_ko": "하단역", "start_en": "Busan Stn", "end_en": "Hadan Stn",
     "via": "남포동,자갈치\nNampo-dong,Jagalchi"},
    {"no": "58-1", "start_ko": "사하구청", "end_ko": "해운대해수욕장", "start_en": "Saha-gu Office",
     "end_en": "Haeundae Beach", "via": "괴정역,대티역,서면\n부전역,연산역\nGoejeong,Daeti,Seomyeon\nBujeon,Yeonsan"},
    {"no": "7-12", "start_ko": "괴정", "end_ko": "다대포", "start_en": "Goejeong", "end_en": "Dadaepo",
     "via": "", "template": "simple_front_192x160"},
]
SELFTEST_STYLES = [{}, {"num_weight": "균일 2픽셀", "arrow_shape": "→"}, {"num_weight": "균일 3픽셀"}]
SELFTEST_EXPECTED = {'360#0': '4dd240bfda9c43c2', '1100#0': 'f55da8885f154f6a', '58-1#0': '15d0ae4361f4ff94', '7-12#0': '44a7a4a1346c97ff', '360#1': 'a30ee2e589538682', '1100#1': 'bdd0df5c2c22033e', '58-1#1': 'bba43390687105c9', '7-12#1': '5fe5c322428ba176', '360#2': '15b0ac3bfbf10f6f', '1100#2': '88e8da31c3aeb8ae', '58-1#2': '1702fe42d4a3cf5b', '7-12#2': '44a7a4a1346c97ff'}  # 검증 완료된 기준 이미지 해시


def selftest_hashes():
    import hashlib
    out = {}
    saved = dict(GLYPHS)
    GLYPHS.clear()
    for si, st in enumerate(SELFTEST_STYLES):
        for r in SELFTEST_ROUTES:
            img = render_route(r, st)
            out[f"{r['no']}#{si}"] = hashlib.sha256(img.tobytes()).hexdigest()[:16]
    GLYPHS.update(saved)
    return out


def run_selftest():
    import PIL, numpy
    print(f"버전: {APP_VERSION}")
    print(f"폰트: {TEXT_FONT}")
    from PIL import features as _ft
    print(f"Pillow {PIL.__version__} / numpy {numpy.__version__} / FreeType {_ft.version('freetype2')} / 배치엔진 BASIC 고정")
    got = selftest_hashes()
    bad = [k for k, v in got.items() if SELFTEST_EXPECTED.get(k) != v]
    for k, v in got.items():
        print(f"  {k:8s} {v}  {'OK' if k not in bad else '불일치 (기준 ' + str(SELFTEST_EXPECTED.get(k)) + ')'}")
    if bad:
        print(f"✖ 자체검사 실패: {len(bad)}/{len(got)}개가 기준 이미지와 다릅니다.")
        return 1
    print(f"✔ 자체검사 통과: {len(got)}개 이미지가 기준과 픽셀 단위로 동일합니다.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        sys.exit(run_selftest())
    if len(sys.argv) > 2 and sys.argv[1] == "--render":  # 테스트용: --render out.bmp
        r = {"no": "360", "start_ko": "서울역 환승센터", "end_ko": "강일IC",
             "start_en": "Seoul Station", "end_en": "Gang-il IC",
             "via": "동묘앞,을지로4가,\n강동역,십자성마을\nDongmyo,Euljuro-4ga,\nGangdong Stn,10JASUNG TOWN"}
        render_route(r).save(sys.argv[2], "BMP")
    else:
        run_gui()
