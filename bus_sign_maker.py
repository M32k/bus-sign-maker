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

APP_VERSION = "2.5.0 (2026-11-26)"
APP_DIR = Path.home() / "BusSignMaker"
OUT_DIR = APP_DIR / "output"
PRESET_FILE = APP_DIR / "presets.json"

# ───────────────────────── 폰트 ─────────────────────────
# 앱에 넣어 둔 폰트를 최우선으로 사용 → 어느 컴퓨터에서든 결과가 픽셀 단위로 같음
BASE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
BUNDLED_FONT = BASE_DIR / "fonts" / "NanumGothic-Regular.ttf"
BUNDLED_BOLD = BASE_DIR / "fonts" / "NanumGothic-Bold.ttf"
CJK_FONT = BASE_DIR / "fonts" / "NotoSansCJKjp-Subset.otf"   # 일본어 가나·한자 (나눔고딕에 없는 글자)
TINY_FONT = BASE_DIR / "fonts" / "Galmuri7-Subset.ttf"       # 7픽셀 픽셀 글꼴 (아주 작은 영문 두 줄용)


def needs_cjk(ch):
    o = ord(ch)
    return (0x3000 <= o <= 0x30FF or 0x31F0 <= o <= 0x31FF or 0x3400 <= o <= 0x4DBF
            or 0x4E00 <= o <= 0x9FFF or 0xF900 <= o <= 0xFAFF)


def _ref(font_path):
    """줄 높이 기준 글자: 한자 글꼴은 '国あAg', 그 외는 '힣갈Ag'"""
    return "国あAg" if font_path and Path(str(font_path)).name.startswith("NotoSansCJK") else "힣갈Ag"
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
    ("old_box", "구형 번호상자", "검은바탕·빨간숫자"),
    ("old_grad", "구형 글자색", "주황 한 색"),
    ("old_arrow", "구형 화살표", "빨간 삼각형"),
    ("new_via", "신형 경유지 화면", "두 줄 크게"),
    ("new_express", "신형 번호색", "흰색"),
    ("fc_num", "풀컬러 번호색", "구분별 색"),
    ("fc_dir", "풀컬러 (방향)색", "초록"),
    ("fc_sep", "풀컬러 구분자", "초록 –"),
]
OLD_BOX_STYLES = ["(기본)", "검은바탕·빨간숫자", "빨강바탕·검은숫자", "빨강바탕·주황숫자"]
OLD_GRAD_STYLES = ["(기본)", "주황 한 색", "그라데이션"]
OLD_ARROW_STYLES = ["(기본)", "빨간 삼각형", "굵은 화살표"]
NEW_VIA_STYLES = ["(기본)", "두 줄 크게", "끔"]
NEW_EXPRESS_STYLES = ["(기본)", "흰색", "구분별 색", "빨강→분홍"]
FC_NUM_STYLES = ["(기본)", "구분별 색", "하늘", "노랑", "자주", "흰색", "초록", "빨강"]
FC_DIR_STYLES = ["(기본)", "초록", "번호색과 같게", "흰색", "노랑"]
FC_SEP_STYLES = ["(기본)", "초록 –", "흰색 ▸"]
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
    out["_raw"] = dict(style or {})
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
    _, rt, _, rb = font.getbbox(_ref(font_path), anchor="la")
    top = int(RENDER_SIZE * 0.4 + rt)
    bot = int(RENDER_SIZE * 0.4 + rb)
    top, bot = min(top, bbox[1]), max(bot, bbox[3])
    big = big.crop((bbox[0], top, bbox[2], bot))
    scale = target_h / big.height
    w = max(1, min(max_w, round(big.width * scale)))
    if stroke is None:
        stroke = CFG["num_stroke"] if is_num else CFG["stroke"]
    if stroke == "bold":
        return uniform_stroke_from_big(big, w, target_h, 1, xbold=True)
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


def uniform_stroke_from_big(big, w, h, stroke, xbold=False):
    """고해상도 글자 → 중심선 추출 → 목표 크기로 축소 → 다시 1px로 → stroke 두께로 칠함.
    모든 획이 같은 픽셀 두께가 됨 (LED 비트맵 폰트처럼)."""
    import numpy as np
    # 획을 칠할 여유(stroke-1)만큼 작은 칸에서 중심선을 뽑아야 오른쪽·아래 끝이 잘리지 않음
    xb = 1 if xbold else 0      # 가로로만 굵게(세로획 2칸, 가로획 1칸): 작은 한글에서도 획이 붙지 않음
    sw, sh_ = max(1, w - stroke + 1 - xb), max(1, h - stroke + 1)
    mid = big.resize((sw * SS, sh_ * SS), Image.BOX)
    a = np.asarray(mid.split()[3]) > 110
    sk = _thin(a)
    # 축소: SS×SS 칸에 중심선이 한 점이라도 지나가면 켬
    low = sk.reshape(sh_, SS, sw, SS).any(axis=(1, 3))
    low = _thin_seq(low)
    if not low.any() and a.any():
        # 마침표처럼 아주 작은 점은 중심선이 사라짐 → 가장 진하게 덮인 칸 하나를 켬
        cov = a.reshape(sh_, SS, sw, SS).sum(axis=(1, 3))
        low = cov == cov.max()
    full = np.zeros((h, w), bool)
    for dy in range(stroke):
        for dx in range(stroke):
            full[dy:dy + sh_, dx:dx + sw] |= low
    if xb:
        # 가로로 한 칸 굵게. 단, 두 칸 옆에 다른 획이 있으면 굵히지 않음
        # (13픽셀 '해'·'대'의 ㅐ 두 줄기가 붙어 덩어리가 되는 것 방지)
        nxt2 = np.zeros_like(low)
        nxt2[:, :-2] = low[:, 2:]
        nxt1 = np.zeros_like(low)
        nxt1[:, :-1] = low[:, 1:]
        grow = low & ~nxt2 & ~nxt1
        for dy in range(stroke):
            full[dy:dy + sh_, stroke:stroke + sw] |= grow
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
    _, rt, _, rb = font.getbbox(_ref(font_path), anchor="la")
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


def _num_em(path, target_h, ref="0123456789"):
    """기준 글자(ref)의 잉크 높이가 target_h가 되는 글꼴 크기"""
    key = (path, target_h, ref)
    if key not in _num_em_cache:
        best = None
        for em10 in range(int(target_h * 10), int(target_h * 20)):
            em = em10 / 10
            f = get_font(path, int(round(em * _NUM_SS)))
            l, t, r, bt = f.getbbox(ref, anchor="ls")
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


def _plain_one(m):
    """'1'의 아래 받침을 지움 (부산 신형 전광판 숫자는 받침 없는 1)"""
    import numpy as np
    h = m.shape[0]
    stem = m[h // 2]
    if not stem.any():
        return m
    m = m.copy()
    for y in range(h - max(2, h // 5), h):
        m[y] &= stem
    xs = np.where(m.any(0))[0]
    return m[:, xs[0]:xs[-1] + 1] if len(xs) else m


def _squeeze_keep_dash(text, cols, track, target_h, max_w, em):
    """글자마다 따로 좁혀서 붙임. '-'는 최소 4칸을 지켜 점(·)처럼 보이지 않게 (부산 판: 128-1)"""
    import numpy as np

    def ink(m):
        xs = np.where(m.any(0))[0]
        return m[:, xs.min():xs.max() + 1] if len(xs) else m[:, :1]
    items = [(ch, None if c is None else ink(c[0])) for ch, c in zip(text, cols)]
    gap = max(1, track + 1)
    dash_w = max(5, target_h // 5)
    sp = round(em * 0.28)
    fixed = gap * (len(items) - 1) + sum(sp for _, m in items if m is None) + sum(dash_w for ch, _ in items if ch == "-")
    var = sum(m.shape[1] for ch, m in items if m is not None and ch != "-")
    k = min(1.0, max(0.3, (max_w - fixed) / max(1, var)))
    parts = []
    for ch, m in items:
        if m is None:
            parts.append(np.zeros((target_h, sp), bool)); continue
        w = dash_w if ch == "-" else max(1, int(m.shape[1] * k))
        parts.append(np.asarray(Image.fromarray(m).resize((w, target_h), Image.NEAREST)))
    out = np.zeros((target_h, sum(p_.shape[1] for p_ in parts) + gap * (len(parts) - 1)), bool)
    x = 0
    for p_ in parts:
        out[:, x:x + p_.shape[1]] = p_
        x += p_.shape[1] + gap
    return out


def render_num_orig(text, color, target_h, max_w, kx=1.0, th=80, track=1, warn=True, plain1=False, keep_dash=False):
    """원본 번호 방식: Arial Bold형 글꼴을 비AA로 찍고, 글자마다 정수 위치, 넘치면 가로 압축."""
    import numpy as np
    path = str(ORIG_NUM_FONT) if ORIG_NUM_FONT.exists() else num_font()
    em = _num_em(path, target_h)
    f = get_font(path, int(round(em * _NUM_SS)))
    asc, desc = f.getmetrics()
    _, rt, _, rb = f.getbbox("0123456789", anchor="ls")
    # 번호에 한글이 섞이면(심야1, 사하10) 번호 글꼴에 한글이 없으므로 나눔고딕 Bold로 같은 높이에 맞춰 그림
    kpath = str(BUNDLED_BOLD) if BUNDLED_BOLD.exists() else text_font()
    kem = _num_em(kpath, target_h, "가힣한") if any(ord(c) > 127 for c in text) else None
    cols = []
    for ch in text:
        if ch == " ":
            cols.append(None); continue
        if ord(ch) > 127 and kem:
            fk = get_font(kpath, int(round(kem * _NUM_SS)))
            ka, kd = fk.getmetrics()
            _, kt, _, kb = fk.getbbox("가힣한", anchor="ls")
            adv = fk.getlength(ch)
            big = Image.new("L", (int(adv + kem * _NUM_SS), ka + kd), 0)
            ImageDraw.Draw(big).text((0, ka), ch, font=fk, fill=255, anchor="ls")
            big = big.crop((0, ka + kt, big.width, ka + kb))
            sm = big.resize((max(1, round(big.width * kx / _NUM_SS)), target_h), Image.BOX)
            cols.append((np.asarray(sm) > th, round(adv * kx / _NUM_SS)))
            continue
        adv = f.getlength(ch)
        ov = glyph_override("NUM", target_h, ch)
        if ov is not None and ov.shape[0] == target_h:
            cols.append((ov, max(ov.shape[1], round(adv * kx / _NUM_SS))))
            continue
        big = Image.new("L", (int(adv + em * _NUM_SS), asc + desc), 0)
        ImageDraw.Draw(big).text((0, asc), ch, font=f, fill=255, anchor="ls")
        big = big.crop((0, asc + rt, big.width, asc + rb))
        sm = big.resize((max(1, round(big.width * kx / _NUM_SS)), target_h), Image.BOX)
        m = np.asarray(sm) > th
        if plain1 and ch == "1":
            m = _plain_one(m)
            cols.append((m, m.shape[1] + max(2, target_h // 8)))
            continue
        cols.append((m, round(adv * kx / _NUM_SS)))
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
    if im.width > max_w and keep_dash and "-" in text:
        out = _squeeze_keep_dash(text, cols, track, target_h, max_w, em)
        if out.shape[1] <= max_w:
            rgba = np.zeros(out.shape + (4,), np.uint8)
            rgba[out, :3] = color
            rgba[out, 3] = 255
            return Image.fromarray(rgba, "RGBA")
        im = Image.fromarray(np.dstack([np.where(out[..., None], np.array(color, np.uint8), 0), (out * 255).astype(np.uint8)]).astype(np.uint8), "RGBA")
    if im.width > max_w:       # 번호는 원본과 같은 최근접 압축이 가장 일치함
        if warn:
            _warn(f"번호 {text}", max_w / im.width, limit=0.62)   # 작은 번호는 원본도 75% 정도로 눌러 그림
        im = im.resize((max_w, im.height), Image.NEAREST)
    return im


def is_wide(ch):
    o = ord(ch)
    return (0xAC00 <= o <= 0xD7A3 or 0x1100 <= o <= 0x11FF or 0x3130 <= o <= 0x318F
            or 0x4E00 <= o <= 0x9FFF or 0xFF00 <= o <= 0xFFEF or 0x3040 <= o <= 0x30FF
            or 0x3400 <= o <= 0x4DBF or 0xF900 <= o <= 0xFAFF)


def render_mono(segs, target_h, full, half, max_w, stroke=None):
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
                fp = str(CJK_FONT) if needs_cjk(ch) and CJK_FONT.exists() else None
                g = render_segments([(ch, color)], target_h, room, font_path=fp, stroke=stroke)
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
        h, fw, hw = SPEC_KO_TWO if korean else SPEC_EN_TWO
        if " " not in text:
            # 띄어쓰기 없는 이름: 작은 글씨 한 줄에 들어가면 한 줄(부산역), 영문은 단어를 자르지 않고 한 줄로 누름
            one = render_mono([(text, color)], h, fw, hw, 10 ** 4)
            if one.width <= bw or not korean:
                put([render_mono([(text, color)], h, fw, hw, bw)], box, [y + (32 - h) // 2])
                return
        a, b2 = split_two(text)
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


# ───────────────────────── 부산 구형 3색 LED (192×32) ─────────────────────────
# 영상(부산 시내버스 37번 구형 LED)을 보고 만든 배치. 빨강·초록·주황(빨강+초록) 3색.
OLD_W, OLD_H = 192, 32
OLD_RED, OLD_GREEN, OLD_AMBER = (255, 40, 30), (40, 255, 60), (255, 185, 0)   # 낮 영상(187번) 기준
OLD_MS = {"big": 2000, "via": 1000, "route": 1500, "slogan": 2000, "logo": 2000, "seats": 2000, "brand": 2000}


# (선택 사항) 스타일 탭 '구형 글자색 = 그라데이션'일 때만 씀. 기본은 주황 한 색
OLD_GRAD = [(0.0, (255, 45, 10)), (0.22, (255, 135, 0)), (0.5, (255, 210, 0)), (0.85, (235, 240, 0)), (1.0, (150, 255, 20))]


def _lerp_stops(stops, t):
    for (t0, c0), (t1, c1) in zip(stops, stops[1:]):
        if t <= t1:
            k = 0 if t1 == t0 else (t - t0) / (t1 - t0)
            return tuple(round(a + (b - a) * k) for a, b in zip(c0, c1))
    return stops[-1][1]


def grad_v(im, stops=OLD_GRAD, y0=None, y1=None):
    """켜진 점을 세로 위치에 따라 색칠 (y0~y1 = 그라데이션 범위, 없으면 글자 잉크 높이)"""
    import numpy as np
    a = np.asarray(im.convert("RGBA")).copy()
    on = a[:, :, 3] > 0
    if not on.any():
        return im
    ys = np.where(on.any(1))[0]
    y0 = ys[0] if y0 is None else y0
    y1 = ys[-1] if y1 is None else y1
    for y in range(a.shape[0]):
        if on[y].any():
            t = min(1.0, max(0.0, (y - y0) / max(1, y1 - y0)))
            a[y, on[y], :3] = _lerp_stops(stops, t)
    return Image.fromarray(a, "RGBA")


def grad_h(im, stops):
    """켜진 점을 가로 위치에 따라 색칠 (신형 급행 번호: 빨강 → 분홍)"""
    import numpy as np
    a = np.asarray(im.convert("RGBA")).copy()
    on = a[:, :, 3] > 0
    if not on.any():
        return im
    xs = np.where(on.any(0))[0]
    for x in range(a.shape[1]):
        if on[:, x].any():
            t = min(1.0, max(0.0, (x - xs[0]) / max(1, xs[-1] - xs[0])))
            a[on[:, x], x, :3] = _lerp_stops(stops, t)
    return Image.fromarray(a, "RGBA")


def _old_colors(st):
    """스타일 탭 색을 고르면 그 색을 쓰고, (기본)이면 3색 LED 기본값"""
    raw = st.get("_raw") or {}

    def pick(key, default):
        v = raw.get(key)
        return COLORS.get(v) if v and v != "(기본)" and COLORS.get(v) else default
    # 실제 구형 3색 LED는 한 색(주황). 그라데이션은 낮 영상의 카메라 색 번짐이라 기본에서 뺌
    grad = (st.get("old_grad") or "주황 한 색") == "그라데이션"
    return {"num": pick("num_big", OLD_AMBER), "box": pick("num_small", OLD_RED), "top": pick("end", OLD_AMBER),
            "via": pick("via", OLD_GREEN), "arrow": pick("arrow", OLD_RED),
            "box_style": st.get("old_box") or "검은바탕·빨간숫자",
            "arrow_style": st.get("old_arrow") or "빨간 삼각형",
            # 직접 고른 색이 있으면 그 색 그대로(그라데이션 안 함)
            "grad_top": grad and raw.get("end") in (None, "", "(기본)"),
            "grad_num": grad and raw.get("num_big") in (None, "", "(기본)")}


def _g(im, c, key="grad_top"):
    return grad_v(im) if c.get(key) else im


def _old_box_w(num):
    """번호 상자 너비: 짧은 번호는 30, 긴 번호(128-1 등)는 사진처럼 상자를 넓혀 숫자를 덜 누름 (최대 50)"""
    if not num:
        return 30
    nat = render_num_orig(num, (255, 255, 255), 24, 10 ** 4, kx=0.96, th=100, track=1, warn=False).width
    if nat <= 26 * 1.3:                     # 2~3자리는 영상(37·187번)과 같은 30칸 상자
        return 30
    return max(30, min(50, round(nat * 0.75) + 4))


def _old_numbox(img, num, c, rule):
    """번호 상자 (초록 테두리). 스타일 탭 '구형 번호상자'로 모양 선택. 구분(급행 등)이면 그 색을 바탕색으로"""
    d = ImageDraw.Draw(img)
    box = c.get("box_style", "검은바탕·빨간숫자")   # 실제 부산 구형: 검은 바탕 + 빨간 숫자 + 초록 테두리
    fill = rule["color"] if rule else c["box"]
    bw = c.get("box_w", 30)
    if box == "검은바탕·빨간숫자":
        d.rectangle([0, 0, bw - 1, 31], outline=OLD_GREEN + (255,))
        n = render_num_orig(num, fill, 24, bw - 4, kx=0.96, th=100, track=1, warn=False, keep_dash=True)
        img.alpha_composite(n, ((bw - n.width) // 2, (32 - n.height) // 2))
        return
    d.rectangle([0, 0, bw - 1, 31], fill=fill + (255,), outline=OLD_GREEN + (255,))
    digit = (0, 0, 0) if box == "빨강바탕·검은숫자" else c["num"]
    n = render_num_orig(num, digit, 24, bw - 4, kx=0.96, th=100, track=1, warn=False, keep_dash=True)
    if digit == (0, 0, 0):          # 꺼진 점으로 숫자 모양을 뚫음 (낮 영상 187번)
        import numpy as np
        a = np.asarray(img).copy()
        m = np.asarray(n)[:, :, 3] > 0
        x0, y0 = (bw - n.width) // 2, (32 - n.height) // 2
        sub = a[y0:y0 + n.height, x0:x0 + n.width]
        sub[m] = (0, 0, 0, 255)
        img.paste(Image.fromarray(a, "RGBA"))
    else:
        img.alpha_composite(n, ((bw - n.width) // 2, (32 - n.height) // 2))


def _old_triangle(img, x, cy, h, color, inner=None):
    """▶ (신형 경유지 구분 등에 씀). inner 색이 있으면 안쪽을 그 색으로"""
    d = ImageDraw.Draw(img)
    w = max(3, h * 2 // 3)
    pts = [(x, cy - h // 2), (x + w, cy), (x, cy + h // 2)]
    d.polygon(pts, fill=color + (255,))
    if inner and h >= 12:
        d.polygon([(x + 2, cy - h // 2 + 3), (x + w - 3, cy), (x + 2, cy + h // 2 - 3)], fill=inner + (255,))
    return w


def _old_arrow(img, x, cy, h, color, style="굵은 화살표"):
    """구형 기점→종점 화살표. 기본 '빨간 삼각형'(37·187번 영상, 128-1·57번 사진),
    '굵은 화살표'는 1번 영상처럼 몸통 + 머리. 너비를 돌려줌"""
    d = ImageDraw.Draw(img)
    if style != "굵은 화살표":
        w = max(4, h * 3 // 4)
        d.polygon([(x + 2, cy - h // 2), (x + 2 + w, cy), (x + 2, cy + h // 2)], fill=color + (255,))
        return w + 4
    head = max(4, h * 5 // 9)               # 머리 길이
    shaft = max(3, h * 2 // 3)              # 몸통 길이
    sh = max(2, h * 4 // 9)                 # 몸통 두께
    d.rectangle([x, cy - sh // 2, x + shaft - 1, cy - sh // 2 + sh - 1], fill=color + (255,))
    d.polygon([(x + shaft - 1, cy - h // 2), (x + shaft - 1 + head, cy), (x + shaft - 1, cy + h // 2)],
              fill=color + (255,))
    return shaft + head


OLD_ROWS = {  # 줄 이름: (줄높이, 한글·한자 칸폭, 반각 칸폭, 획)
    "top": (16, 15, 9, "bold"),      # 경유지 화면 윗줄 (기점 ▶ 종점)
    "via": (12, 14, 8, "bold"),      # 경유지 화면 아랫줄 (영문)
    "via_ko": (14, 14, 8, 1),        # 경유지 화면 아랫줄 (한글: 12픽셀이면 ㅇ이 ㅁ처럼 보여 14)
    "big": (20, 19, 11, "bold"),     # 행선 화면 한 줄
    "two": (13, 14, 9, "bold"),      # 행선 화면 두 줄
    "ntop": (17, 16, 9, "bold"),     # 신형: 윗줄 'OO행' (영상 99번: 글자가 판 위쪽 절반 이상을 채움)
    "ndest": (24, 24, 12, 2),        # 신형 번호 + 'OO행' (1000번 영상: 글자 높이 ≈ 번호의 80%)
    "ndest2": (22, 22, 11, 2),       # 위가 안 들어갈 때
    "nvia2": (15, 15, 7, 1),         # 신형 경유지 두 줄 화면 (15픽셀이어야 '진·정'의 ㅈ이 또렷함)
    "nvia": (14, 14, 8, "bold"),     # 신형: 아랫줄 경유지 (13픽셀이면 '청'의 ㅇ이 ㅁ처럼 보여서 14)
    "seat": (22, 19, 12, "bold"),
    "nbig": (20, 19, 11, 2),         # 신형 한 줄 'OO행'·문구 (영상: 2픽셀 굵은 획)
    "nsub": (16, 16, 9, 2),          # 신형 윗줄 '다 대 포 행' (벌려 씀)    # 잔여좌석 글자
}


def _old_text(text, color, row, max_w):
    h, full, half, stroke = OLD_ROWS[row]
    # 작은 줄의 한글·한자는 굵게 하면 획이 뭉개짐(장안산단, 公營) → 그 줄만 가늘게
    if (h <= 13 or row == "nvia") and (any(needs_cjk(ch) for ch in text) or (row in ("via", "nvia") and has_hangul(text))):
        stroke = 1
    return render_mono([(text, color)], h, full, half, max_w, stroke=stroke)


def _smart_split(text):
    """두 줄로 나누기: 띄어쓰기 → 영문/한자 경계(Gowon|アパート) → 글자 절반 순"""
    if " " in text:
        return split_two(text)
    bounds = [i for i in range(1, len(text)) if is_wide(text[i]) != is_wide(text[i - 1])]
    if bounds:
        i = min(bounds, key=lambda b: abs(b - len(text) / 2))
        return text[:i], text[i:]
    h = (len(text) + 1) // 2
    return text[:h], text[h:]


def _old_top_row(img, start, end, c):
    """윗줄: 기점 ▶ 종점 (주황, 굵게). 칸이 모자라면 기점·종점을 같은 비율로 좁혀 다시 그림"""
    x0, x1, h, y = c.get("box_w", 30) + 2, 192, 16, 1
    tri = 14
    gap = 3
    a = _old_text(start, c["top"], "top", 10 ** 4) if start else None
    b = _old_text(end, c["top"], "top", 10 ** 4) if end else None
    room = (x1 - x0) - (tri + 2 * gap if a and b else 0)
    wa, wb = (a.width if a else 0), (b.width if b else 0)
    if wa + wb > room:
        k = room / (wa + wb)
        _warn(f"{start}▶{end}", k)
        if a:
            a = _old_text(start, c["top"], "top", max(4, int(wa * k)))
        if b:
            b = _old_text(end, c["top"], "top", max(4, int(wb * k)))
    total = (a.width if a else 0) + (b.width if b else 0) + (tri + 2 * gap if a and b else 0)
    x = x0 + (x1 - x0 - total) // 2
    if a:
        img.alpha_composite(_g(a, c), (x, y)); x += a.width + gap
    if a and b:
        _old_arrow(img, x, y + h // 2, 11, c["arrow"], c["arrow_style"]); x += tri + gap
    if b:
        img.alpha_composite(_g(b, c), (x, y))


def _old_en_parts(text, c):
    """영문 경유지가 비례폭 12칸으로도 판(약 160칸)에 85% 이상 눌러야 하면 단어 경계에서 두 화면으로 나눔"""
    W = 192 - (c.get("box_w", 30) + 2)
    nat = render_segments([(text, (255, 255, 255))], 12, 10 ** 4, font_path=str(LATIN_BOLD), stroke="bold").width
    if nat * 0.85 <= W or " " not in text:
        return [text]
    a, b = split_two(text)
    return [a.strip(), b.strip()]


def _old_bottom_row(img, text, c):
    """아랫줄: 경유지 하나 (초록, 가는 획), 가운데 정렬"""
    X0 = c.get("box_w", 30) + 2
    W = 192 - X0
    one = _old_text(text, c["via"], "via", 10 ** 4)
    if one.width > W and not has_hangul(text) and not any(needs_cjk(ch) for ch in text):
        # 긴 영문: 고정폭 대신 비례폭 굵은 글씨(12칸). 7픽셀은 승객이 못 읽어서 쓰지 않음
        t = render_segments([(text, c["via"])], 12, W, font_path=str(LATIN_BOLD), stroke="bold")
        img.alpha_composite(t, (X0 + (W - t.width) // 2, 20))
        return
    if has_hangul(text):                    # 사진(128-1번 '활 천 고 개')처럼 글자 사이를 벌림
        t = _spaced(text, c["via"], "via_ko", W - 10)
        img.alpha_composite(t, (X0 + (W - t.width) // 2, 18))
        return
    t = _old_text(text, c["via"], "via", W)
    img.alpha_composite(t, (X0 + (W - t.width) // 2, 20))


def _old_route_page(img, start, end, c, breaks=(None, None)):
    """행선 화면: 기점(두 줄 가능) ▶ 종점(크게). breaks = 사용자가 '|'로 정한 두 줄 나눔"""
    def block(text, x0, x1, forced=None):
        bw = x1 - x0
        if not text:
            return
        if forced:                          # '김해|구산동' → 김해 / 구산동
            for i, part in enumerate(forced):
                li = _old_text(part, c["top"], "two", bw)
                img.alpha_composite(_g(li, c), (x0 + (bw - li.width) // 2, 2 + i * 15))
            return
        one = _old_text(text, c["top"], "big", 10 ** 4)
        latin_word = " " not in text and not any(is_wide(ch) for ch in text)
        mixed = any(is_wide(ch) for ch in text) and any(ch.isascii() and ch.isalpha() for ch in text)
        if one.width <= bw or latin_word or (one.width * 0.75 <= bw and not mixed):
            # 들어가면 한 줄. 한 글자 체계만 쓰면 25%까지 눌러서 한 줄(영상의 '고원아파트'),
            # 영문+한자처럼 섞여 있으면 경계에서 두 줄(Gowon / 公寓)
            one = _old_text(text, c["top"], "big", bw)
            img.alpha_composite(_g(one, c), (x0 + (bw - one.width) // 2, 6))
            return
        for i, part in enumerate(_smart_split(text)):
            li = _old_text(part.strip(), c["top"], "two", bw)
            img.alpha_composite(_g(li, c), (x0 + (bw - li.width) // 2, 2 + i * 15))
    x0 = c.get("box_w", 30) + 2
    mid = x0 + (192 - x0 - 16) // 2
    block(start, x0, mid, breaks[0])
    if start and end:
        _old_arrow(img, mid + 1, 16, 14, c["arrow"], c["arrow_style"])
    block(end, mid + 16, 192, breaks[1])


def logo_to_led(path, w, h, palette=None, full=False):
    """사용자가 고른 이미지 → 판 크기에 맞춰 줄이고 3색 LED 색(빨강·초록·주황)으로 바꿈.
    full=True 이면 풀컬러 판(부산 신형)용: 색은 그대로 두고 어두운 점만 끔"""
    import numpy as np
    pal = palette or [(0, 0, 0), OLD_RED, OLD_GREEN, OLD_AMBER]
    src = Image.open(path).convert("RGBA")
    bg = Image.new("RGBA", src.size, (0, 0, 0, 255))
    src = Image.alpha_composite(bg, src).convert("RGB")
    src.thumbnail((w, h), Image.LANCZOS)
    a = np.asarray(src).astype(int)
    if full:
        q = np.where(a.max(-1, keepdims=True) < 60, 0, a).astype(np.uint8)
        out = Image.new("RGBA", (w, h), (0, 0, 0, 255))
        out.alpha_composite(Image.fromarray(q, "RGB").convert("RGBA"), ((w - src.width) // 2, (h - src.height) // 2))
        return out
    P = np.array(pal)
    d = ((a[:, :, None, :] - P[None, None, :, :]) ** 2).sum(-1)
    q = P[d.argmin(-1)].astype(np.uint8)
    out = Image.new("RGBA", (w, h), (0, 0, 0, 255))
    out.alpha_composite(Image.fromarray(q, "RGB").convert("RGBA"), ((w - src.width) // 2, (h - src.height) // 2))
    return out


def _flat_stops(lines):
    return [s for line in lines for s in line]


def tpl_busan_old(r, st):
    """부산 구형 3색 LED 192×32 — 여러 장(페이지)을 돌려줌. 각 장의 info['duration'] = 표시 시간(ms)"""
    num = str(r.get("no", "")).strip()
    c = _old_colors(st)
    c["box_w"] = _old_box_w(str(r.get("no", "")).strip())
    rule = kind_style(r.get("kind"), st)
    pages = []

    def new():
        return Image.new("RGBA", (OLD_W, OLD_H), (0, 0, 0, 255))

    def add(img, kind):
        im = img.convert("RGB")
        im.info["duration"] = OLD_MS[kind]
        im.info["page_kind"] = kind
        pages.append(im)

    def big():
        img = new()
        if num:
            color = rule["color"] if rule and not rule["invert"] else c["num"]
            n = render_num_orig(num, color, 28, OLD_W - 8, kx=1.0, th=100, track=2)
            if rule and rule["invert"]:
                n = invert_box(n, rule["color"])
            elif not rule:
                n = _g(n, c, "grad_num")
            img.alpha_composite(n, ((OLD_W - n.width) // 2, (OLD_H - n.height) // 2))
        add(img, "big")

    ko_lines, en_lines = parse_via(r.get("via", ""))
    ko, en = _flat_stops(ko_lines), _flat_stops(en_lines)
    start, end = r.get("start_ko", ""), r.get("end_ko", "")
    big()
    for i, stop in enumerate(ko):
        texts = [stop]
        if i < len(en) and en[i]:
            texts += _old_en_parts(en[i], c)        # 긴 영문은 7픽셀로 줄이지 않고 두 화면으로 나눔
        for text in texts:
            img = new()
            _old_numbox(img, num, c, rule)
            _old_top_row(img, start, end, c)
            _old_bottom_row(img, text, c)
            add(img, "via")
        if i % 2 == 1 and i + 1 < len(ko):      # 정류장 두 개마다 큰 번호 한 번
            big()
    logo = (r.get("logo") or "").strip()
    if logo:
        lp = Path(logo).expanduser()
        if lp.exists():
            add(logo_to_led(lp, OLD_W, OLD_H), "logo")
        else:
            WARN.append((f"로고 파일 없음: {logo}", 0))
    if r.get("slogan", "").strip():
        img = new()
        t = _old_text(r["slogan"].strip(), c["top"], "big", OLD_W - 8)
        img.alpha_composite(_g(t, c), ((OLD_W - t.width) // 2, 6))
        add(img, "slogan")
    br = r.get("_br") or {}
    for sk, ek in (("start_ko", "end_ko"), ("start_en", "end_en"), ("start_ja", "end_ja"), ("start_zh", "end_zh")):
        a, b = r.get(sk, "").strip(), r.get(ek, "").strip()
        fa, fb = br.get(sk), br.get(ek)
        if sk == "start_en" and not (a or b):        # 영문이 비어 있으면 한글에서 자동 변환
            def auto(k):
                if br.get(k):                        # '김해|구산동' → Gimhae / Gusan-dong
                    return tuple(romanize_name(x) for x in br[k])
                v = r.get(k, "").strip()
                return romanize_name(v) if v else ""
            a, b = auto("start_ko"), auto("end_ko")
            if isinstance(a, tuple):
                fa, a = a, " ".join(a)
            if isinstance(b, tuple):
                fb, b = b, " ".join(b)
        if sk in ("start_ja", "start_zh") and not (a or b) and not st.get("old_no_auto"):
            # 일본어·중국어도 비어 있으면 자동 (사전 → 일본어는 가타카나). '김해|구산동'은 나눈 그대로 번역
            lang = sk[-2:]

            def tr(k):
                if br.get(k):
                    parts = tuple(translate_place(x, lang) for x in br[k])
                    return parts if all(parts) else ""
                v = r.get(k, "").strip()
                return translate_place(v, lang) if v else ""
            ta, tb = tr("start_ko"), tr("end_ko")
            need_a, need_b = bool(r.get("start_ko", "").strip()), bool(r.get("end_ko", "").strip())
            if (need_a and not ta) or (need_b and not tb):
                if lang == "zh" and (need_a or need_b):
                    WARN.append((f"중문 행선을 몰라서 생략: {r.get('start_ko','')}▶{r.get('end_ko','')} → 기점(중)·종점(중) 칸에 적거나 번역 사전에 추가하세요", 0))
            else:
                a, b = ta, tb
                if isinstance(a, tuple):
                    fa, a = a, "".join(a)
                if isinstance(b, tuple):
                    fb, b = b, "".join(b)
        if a or b:
            img = new()
            _old_numbox(img, num, c, rule)
            _old_route_page(img, a, b, c, (fa, fb))
            add(img, "route")
    seats = seats_text(r)
    if seats:
        add(seats_page(seats, OLD_RED, OLD_AMBER), "seats")
    return pages


def seats_text(r):
    """잔여좌석 칸: 숫자만 남김 (비어 있으면 화면 생략)"""
    v = str(r.get("seats", "") or "").strip()
    if v.endswith(".0"):
        v = v[:-2]
    v = "".join(ch for ch in v if ch.isdigit())
    return v[:3]


def seats_page(seats, label_color, num_color, w=OLD_W, h=OLD_H):
    """잔여좌석 화면: 왼쪽 '잔여좌석'(굵게) + 오른쪽 큰 숫자 (경기·부산 좌석버스 영상의 배치)"""
    img = Image.new("RGBA", (w, h), (0, 0, 0, 255))
    lab = _old_text("잔여좌석", label_color, "seat", w // 2)
    n = render_num_orig(seats, num_color, 28, w // 2 - 16, kx=0.85, th=100, track=3, warn=False)
    gap = 14                                    # 영상: 글자와 숫자가 가까이 붙어 가운데 모임
    x = (w - (lab.width + gap + n.width)) // 2
    img.alpha_composite(lab, (x, (h - lab.height) // 2))
    img.alpha_composite(n, (x + lab.width + gap, (h - n.height) // 2))
    return img


# ───────────────────────── 부산 신형 풀컬러 LED (192×32) ─────────────────────────
# 영상(부산 2029번 전면 신형 전광판, 99번 측면)을 보고 만든 배치.
#  화면 순서: [구분+번호] → [번호 + 'OO행' 한 줄] → [번호 + 'OO행' + 작은 경유지] … → 영문 → 잔여좌석 → 문구
#  번호색: 일반 = 청록빛 하늘색(99번), 급행·좌석 = 빨강→분홍 가로 그라데이션(2029번)
#  화면이 바뀔 때는 효과 없이 바로 바뀜
NEW_NUM = (0, 225, 235)        # '구분별 색'일 때 일반 번호 (99번 측면 영상)
NEW_NUM_WHITE = (255, 255, 255)
NEW_TEXT = (250, 245, 255)
NEW_ARROW = (255, 60, 200)
NEW_YELLOW = (255, 215, 0)
NEW_PINK = (255, 50, 220)
NEW_VIA_ARROW = (0, 170, 255)   # 경유지 사이 ▶ : 측면 신형 영상처럼 파랑
LATIN_BOLD = BASE_DIR / "fonts" / "LiberationSans-Bold.ttf"       # 'Busan is good' (1000번·99번 영상)
NEW_RED_PINK = [(0.0, (255, 40, 40)), (1.0, (255, 40, 220))]
NEW_KIND_COLORS = {"심야": (60, 120, 255), "마을": (60, 230, 60), "순환": (255, 150, 0)}


def _new_colors(st):
    raw = st.get("_raw") or {}

    def pick(key, default):
        v = raw.get(key)
        return COLORS.get(v) if v and v != "(기본)" and COLORS.get(v) else default
    return {"num": pick("num_big", NEW_NUM), "top": pick("end", NEW_TEXT), "via": pick("via", NEW_TEXT),
            "arrow": pick("arrow", NEW_ARROW), "num_custom": raw.get("num_big") not in (None, "", "(기본)"),
            "express": st.get("new_express") or "흰색"}


def _num_fixed_gap(num, color, h, max_w, gap=1):
    """숫자를 한 자씩 그려서 글자 사이를 정확히 gap칸으로 붙임 (신형 번호: 1칸).
    넘치면 원래 방식(최근접 압축)으로 판 폭에 맞춤"""
    import numpy as np
    parts = []
    for ch in num:
        if ch == " ":
            parts.append(np.zeros((h, max(2, h // 6)), bool)); continue
        g = render_num_orig(ch, color, h, 10 ** 4, kx=0.95, th=100, track=0, warn=False)
        m = np.asarray(g)[:, :, 3] > 0
        xs = np.where(m.any(0))[0]
        if len(xs):
            m = m[:, xs.min():xs.max() + 1]
        if ch == "-":                                   # 붙임표는 숫자 가운데 높이로 최소 폭 유지
            m = np.asarray(Image.fromarray(m).resize((max(m.shape[1], h // 3), h), Image.NEAREST))
        parts.append(m)
    W = sum(p_.shape[1] for p_ in parts) + gap * (len(parts) - 1)
    out = np.zeros((h, max(1, W)), bool)
    x = 0
    for p_ in parts:
        out[:, x:x + p_.shape[1]] = p_
        x += p_.shape[1] + gap
    rgba = np.zeros(out.shape + (4,), np.uint8)
    rgba[out, :3] = color
    rgba[out, 3] = 255
    im = Image.fromarray(rgba, "RGBA")
    if im.width > max_w:
        return render_num_orig(num, color, h, max_w, kx=0.95, th=100, track=-2, warn=False, keep_dash=True)
    return im


def _new_num_paint(num, c, kind, st, h=28, max_w=92, track=-2):
    """번호 그리기 + 색 (기본 흰색). 사용자가 확인한 사진 모양: 높이 28, 굵은 숫자, 숫자 사이 1칸"""
    n = _num_fixed_gap(num, (255, 255, 255), h, max_w)
    custom = (st.get("kinds") or {}).get(kind)
    if custom and (custom.get("color", "(기본)"), bool(custom.get("invert"))) != KIND_DEFAULTS.get(kind, ("(기본)", False)):
        # 스타일 탭에서 구분 색을 기본값과 다르게 바꿨으면 그것
        rule = kind_style(kind, st)
        if rule:
            n = render_num_orig(num, rule["color"], h, max_w, kx=0.95, th=100, track=track, warn=False)
            return invert_box(n, rule["color"]) if rule["invert"] else n
    if c["num_custom"]:
        return recolor(n, c["num"])
    mode = c["express"]
    if mode == "빨강→분홍":
        return grad_h(n, NEW_RED_PINK)
    if mode == "구분별 색":
        if kind in ("급행", "좌석"):
            return grad_h(n, NEW_RED_PINK)
        return recolor(n, NEW_KIND_COLORS.get(kind, NEW_NUM))
    return recolor(n, NEW_NUM_WHITE)                    # 기본: 흰색 번호


def recolor(im, color):
    import numpy as np
    a = np.asarray(im.convert("RGBA")).copy()
    a[a[:, :, 3] > 0, :3] = color
    return Image.fromarray(a, "RGBA")


def _dest(end, lang="ko"):
    end = (end or "").strip()
    if not end:
        return ""
    if lang == "ko":
        return end if end.endswith("행") else end + "행"
    return end if end.lower().startswith(("to ", "for ")) else "To " + end


def _tiny_chain(stops, color, sep_color, max_w):
    """작은 글씨(7픽셀) 경유지 줄: 경유지 ▸ 경유지 ▸ 경유지 (▸는 점으로 직접 그림)"""
    ims = [render_segments_pixel([(s_, color)], 7, 10 ** 4, str(TINY_FONT)) for s_ in stops]
    sep_w, gap = 3, 3
    w = sum(i.width for i in ims) + (len(ims) - 1) * (sep_w + 2 * gap)
    out = Image.new("RGBA", (max(1, w), 7), (0, 0, 0, 0))
    d = ImageDraw.Draw(out)
    x = 0
    for k, im in enumerate(ims):
        out.alpha_composite(im, (x, 0))
        x += im.width
        if k < len(ims) - 1:
            x += gap
            d.polygon([(x, 1), (x + sep_w - 1, 3), (x, 5)], fill=sep_color + (255,))
            x += sep_w + gap
    return out if out.width <= max_w else squeeze_w(out, max_w)


def _spaced(text, color, row, room):
    """글자를 한 자씩 그려 칸에 고르게 벌려 놓음 (영상: '다 대 포 행')"""
    chars = [ch for ch in text if not ch.isspace()]
    if not chars:
        return Image.new("RGBA", (1, 1), (0, 0, 0, 0))
    ims = [_old_text(ch, color, row, 10 ** 4) for ch in chars]
    tot = sum(i.width for i in ims)
    if tot > room or len(ims) == 1 or not has_hangul(text):
        return _old_text(text, color, row, room)
    avg = tot / len(ims)
    gap = int(min((room - tot) / (len(ims) - 1), avg * 0.55))
    h = max(i.height for i in ims)
    out = Image.new("RGBA", (tot + gap * (len(ims) - 1), h), (0, 0, 0, 0))
    x = 0
    for im in ims:
        out.alpha_composite(im, (x, h - im.height))
        x += im.width + gap
    return out


def tpl_busan_new(r, st):
    """부산 신형 풀컬러 LED 192×32 — 여러 장. info['trans'] = 'cut' (효과 없이 바로 바뀜)
    영상(부산 1000번 급행 전면, 2029번, 99번 측면) 순서:
      [번호만] → [번호 + 'OO행' 한 줄] → [번호 + 'O O O 행' + 영문] → [… + 중문/일문] → (경유지) → 잔여좌석 → 문구"""
    num = str(r.get("no", "")).strip()
    kind = (r.get("kind") or "").strip()
    c = _new_colors(st)
    pages = []

    def new():
        return Image.new("RGBA", (OLD_W, OLD_H), (0, 0, 0, 255))

    def add(img, k, ms=None):
        im = img.convert("RGB")
        im.info["duration"] = ms or OLD_MS[k]
        im.info["page_kind"] = k
        im.info["trans"] = "cut"       # 신형: 효과 없이 바로 다음 화면으로 (사용자 요청)
        pages.append(im)

    n_img = _new_num_paint(num, c, kind, st) if num else None
    x0 = (n_img.width + 6) if n_img else 2
    room = OLD_W - 1 - x0

    def with_num():
        img = new()
        if n_img:
            img.alpha_composite(n_img, (1, (OLD_H - n_img.height) // 2))
        return img

    def dest_page(dest):
        """번호 + 'OO행' 한 줄. 1000번 영상처럼 글자 높이가 번호의 80%쯤(24칸), 획 2칸, 글자 사이를 벌림.
        길어서 안 들어가면 22칸 → 예전 크기(20칸) 순으로 줄임"""
        img = with_num()
        t = None
        cands = [(dest, row) for row in ("ndest", "ndest2")]
        if dest.endswith("행") and len(dest) > 2:       # 길면 실제 전광판처럼 '행'을 빼고 크게
            cands += [(dest[:-1], row) for row in ("ndest", "ndest2")]
        for dest, row in cands:
            if has_hangul(dest) or any(needs_cjk(ch) for ch in dest):
                ims = [_old_text(ch, c["top"], row, 10 ** 4) for ch in dest if not ch.isspace()]
                tot = sum(i.width for i in ims)
                if tot <= room:
                    gap = max(1, min(round(tot / len(ims) * 0.35), (room - tot) // max(1, len(ims) - 1))) if len(ims) > 1 else 0
                    hh = max(i.height for i in ims)
                    t = Image.new("RGBA", (tot + gap * (len(ims) - 1), hh), (0, 0, 0, 0))
                    xx = 0
                    for im in ims:
                        t.alpha_composite(im, (xx, hh - im.height))
                        xx += im.width + gap
                    break
            else:
                one = _old_text(dest, c["top"], row, 10 ** 4)
                if one.width <= room:
                    t = one
                    break
        if t is None:
            t = _old_text(dest, c["top"], "big", room)      # dest = '행'을 뺀 이름 (그래도 길면 좁혀서)
        img.alpha_composite(t, (x0 + (room - t.width) // 2, (OLD_H - t.height) // 2))
        add(img, "route", 1000)

    def sub_page(dest, sub, ms=1200):
        """윗줄 'O O O 행'(벌려 씀) + 아랫줄 번역(영문·중문·일문)"""
        img = with_num()
        rm, xs = room, x0
        if not (has_hangul(sub) or any(needs_cjk(ch) for ch in sub)):
            nat = render_segments([(sub, c["top"])], 12, 10 ** 4, font_path=str(LATIN_BOLD), stroke="bold").width
            if nat * 0.87 > room:                         # 영문이 너무 길면 번호를 빼고 판 전체 폭으로
                img, rm, xs = new(), OLD_W - 2, 1
        room_, x0_ = rm, xs
        if dest.endswith("행") and _old_text(dest, c["top"], "ntop", 10 ** 4).width > room_:
            dest = dest[:-1]                              # 길면 '행'을 뺌
        t = _spaced(dest, c["top"], "ntop", room_)
        img.alpha_composite(t, (x0_ + (room_ - t.width) // 2, 0))
        v = _old_text(sub, c["top"], "via", 10 ** 4)
        if v.width > room_ and not (has_hangul(sub) or any(needs_cjk(ch) for ch in sub)):
            # 긴 영문은 고정폭으로 누르면 못 읽으니 비례폭 굵은 글씨(12칸)로
            v = render_segments([(sub, c["top"])], 12, room_, font_path=str(LATIN_BOLD), stroke="bold")
        elif v.width > room_:
            v = _old_text(sub, c["top"], "via", room_)
        img.alpha_composite(v, (x0_ + (room_ - v.width) // 2, 19))
        add(img, "route", ms)

    def via_groups(stops):
        out, cur = [], []
        for s_ in stops:
            if cur and _tiny_chain(cur + [s_], c["via"], c["arrow"], 10 ** 4).width > room:
                out.append(cur)
                cur = []
            cur.append(s_)
        if cur:
            out.append(cur)
        return out

    def via_lines(stops):
        """경유지를 판 전체 폭(190칸)에 '경유지 ▶ 경유지' 줄로 채움. 두 줄씩 한 화면 → 5개 이상 보임"""
        tri, gap, W_ = 6, 3, OLD_W - 2
        ims = [_old_text(s_, c["via"], "nvia2", W_) for s_ in stops]
        lines, cur, w = [], [], 0
        for im in ims:
            add_w = im.width + (tri + 2 * gap if cur else 0)
            if cur and w + add_w > W_:
                lines.append(cur)
                cur, w, add_w = [], 0, im.width
            cur.append(im)
            w += add_w
        if cur:
            lines.append(cur)
        out = []
        for ln in lines:
            lw = sum(i.width for i in ln) + (len(ln) - 1) * (tri + 2 * gap)
            li = Image.new("RGBA", (lw, 15), (0, 0, 0, 0))
            x = 0
            for k, im in enumerate(ln):
                li.alpha_composite(im, (x, 15 - im.height))
                x += im.width
                if k < len(ln) - 1:
                    _old_triangle(li, x + gap, 7, 9, NEW_VIA_ARROW)
                    x += tri + 2 * gap
            out.append(li)
        return out

    def via_full_pages(stops):
        """측면 신형 영상(128-1번)처럼 흰 글자 + 파란 ▶. 두 줄 = 한 화면"""
        lines = via_lines(stops)
        for i in range(0, len(lines), 2):
            img = new()
            pair = lines[i:i + 2]
            ys = [0, 17] if len(pair) == 2 else [8]
            for li, y in zip(pair, ys):
                img.alpha_composite(li, ((OLD_W - li.width) // 2, y))
            add(img, "via", 2500)

    def via_page(dest, stops):
        img = with_num()
        t = _spaced(dest, c["top"], "ntop", room)
        img.alpha_composite(t, (x0 + (room - t.width) // 2, 1))
        if TINY_FONT.exists():
            v = _tiny_chain(stops, c["via"], c["arrow"], room)
            img.alpha_composite(v, (x0 + (room - v.width) // 2, 23))
        add(img, "via")

    if num:                                              # 번호만 크게 (멀리서도 보이게)
        img = new()
        n = _new_num_paint(num, c, kind, st, h=30, max_w=OLD_W - 8, track=-2)
        img.alpha_composite(n, ((OLD_W - n.width) // 2, (OLD_H - n.height) // 2))
        add(img, "big")
    ko_lines, en_lines = parse_via(r.get("via", ""))
    ko = _flat_stops(ko_lines)
    end_ko = (r.get("end_ko") or "").strip()
    dest_ko = _dest(end_ko)
    if dest_ko:
        dest_page(dest_ko)
        end_en = (r.get("end_en") or "").strip() or romanize_name(end_ko)
        end_zh = (r.get("end_zh") or "").strip() or translate_place(end_ko, "zh")
        end_ja = (r.get("end_ja") or "").strip() or translate_place(end_ko, "ja")
        if not end_zh:
            WARN.append((f"중문 행선지를 몰라서 생략: {end_ko} → 종점(중) 칸에 적거나 번역 사전에 추가하세요", 0))
        shown = set()
        for sub in (end_en, end_zh, end_ja):              # 다대포행 밑에 영어 → 중국어 → 일본어
            if sub and sub not in shown:                  # 중국어·일본어 한자가 같으면(多大浦) 한 번만
                shown.add(sub)
                sub_page(dest_ko, sub)
        if st.get("new_via") != "끔" and ko:            # 7픽셀 작은 경유지는 승객이 못 읽어서 없앰          # 기본: 경유지만 두 줄, 한 화면에 5개 이상
            via_full_pages(ko)
            en = _flat_stops(en_lines)
            if en:
                via_full_pages(en)                       # 영문 경유지도 같은 방식
    elif r.get("end_en"):
        dest_page(_dest(r["end_en"], "en"))
    seats = seats_text(r)
    if seats:                                            # 신형 잔여좌석: 글자·숫자 모두 흰색
        add(seats_page(seats, NEW_TEXT, NEW_NUM_WHITE), "seats")
    logo = (r.get("logo") or "").strip()
    if logo:
        lp = Path(logo).expanduser()
        if lp.exists():
            add(logo_to_led(lp, OLD_W, OLD_H, full=True), "logo")
        else:
            WARN.append((f"로고 파일 없음: {logo}", 0))
    if r.get("slogan", "").strip():                      # 'Busan is good' : 분홍
        img = new()
        t = recolor(_old_text(r["slogan"].strip(), NEW_TEXT, "big", OLD_W - 8), NEW_PINK)
        img.alpha_composite(t, ((OLD_W - t.width) // 2, (OLD_H - t.height) // 2))
        add(img, "slogan", 2500)
    return pages


# ───────────────────────── 풀컬러 방향 표시 (192×32) ─────────────────────────
# 사용자가 보내 준 풀컬러 LED 행선판 사진들(148-1 부산대학교 방향, 864 인경리–청주의료원, 31 의정부역 방향 등)의 배치.
#  [번호 + 행선지 + (방 향)]  →  [번호 + 기점 – 종점]  →  [번호 + 영문 기점 – 종점]  →  [번호만 크게]
#  번호: 굵은 숫자, 구분별 색(일반 하늘·급행 자주·좌석 빨강·마을 노랑…). 행선지는 흰색, '(방 향)'은 초록.
FC_COLORS = {"하늘": (80, 225, 255), "노랑": (255, 235, 40), "자주": (235, 60, 235), "흰색": (255, 255, 255),
             "초록": (60, 235, 80), "빨강": (255, 70, 60)}
FC_KIND = {"일반": "하늘", "급행": "자주", "좌석": "빨강", "마을": "노랑", "심야": "노랑", "순환": "초록"}


def _fc_text_fit(text, color, rows, room, strict=False):
    """여러 크기(rows) 중 칸(room)에 85% 이상 눌리지 않고 들어가는 가장 큰 것으로. 영문은 비례폭 굵은 글씨.
    strict=True 이면 안 들어갈 때 None (두 줄로 나눌지 정할 때)"""
    latin = not (has_hangul(text) or any(needs_cjk(ch) for ch in text))
    for row in rows:
        if latin:
            h = OLD_ROWS[row][0]
            im = render_segments([(text, color)], max(9, h - 3), 10 ** 4, font_path=str(LATIN_BOLD), stroke="bold")
        else:
            im = _old_text(text, color, row, 10 ** 4)
        if im.width * 0.85 <= room:
            if im.width > room:
                im = squeeze_w(im, room)
            return im
    if strict:
        return None
    if latin:
        return render_segments([(text, color)], 9, room, font_path=str(LATIN_BOLD), stroke="bold")
    return _old_text(text, color, rows[-1], room)


def _fc_block(text, color, x0, x1, img, forced=None):
    """기점·종점 한 덩어리: 한 줄로 크게, 안 되면 두 줄 (864 '청주 / 의료원', 842-1 영문)"""
    bw = x1 - x0
    if not text:
        return
    if not forced:
        one = _fc_text_fit(text, color, ("big", "top"), bw, strict=True)
        if one is not None:
            img.alpha_composite(one, (x0 + (bw - one.width) // 2, (OLD_H - one.height) // 2))
            return
        latin1 = " " not in text and not has_hangul(text) and not any(needs_cjk(ch) for ch in text)
        if len(text) <= 4 or latin1:
            one = _fc_text_fit(text, color, ("two",), bw, strict=True)
            if one is not None and one.width <= bw:
                img.alpha_composite(one, (x0 + (bw - one.width) // 2, (OLD_H - one.height) // 2))
                return
        if latin1 and len(text) >= 8:                  # 긴 영문 한 단어는 가운데서 끊어 두 줄 (하이픈)
            cut = text.find("-", len(text) // 3)
            cut = cut + 1 if 0 < cut < len(text) * 2 // 3 else len(text) // 2
            forced = (text[:cut] if text[:cut].endswith("-") else text[:cut] + "-", text[cut:])
        elif len(text) <= 2:
            one = _fc_text_fit(text, color, ("two", "via"), bw)
            img.alpha_composite(one, (x0 + (bw - one.width) // 2, (OLD_H - one.height) // 2))
            return
    parts = list(forced) if forced else [p_.strip() for p_ in _smart_split(text)]
    if len(parts) == 1 and " " in text:
        i = text.rfind(" ", 0, len(text) // 2 + 3)
        parts = [text[:i], text[i + 1:]] if i > 0 else parts
    if len(parts) == 1 and len(text) >= 5:            # 한 덩어리 이름도 사진처럼 반으로 ('청주 / 의료원')
        h = (len(text) + 1) // 2
        parts = [text[:h], text[h:]]
    ims = [_fc_text_fit(pt, color, ("two", "via"), bw) for pt in parts[:2]]
    y = 2 if len(ims) == 2 else (OLD_H - ims[0].height) // 2
    for im in ims:
        img.alpha_composite(im, (x0 + (bw - im.width) // 2, y))
        y += 15


def tpl_fullcolor_dir(r, st):
    """풀컬러 방향 표시 192×32 — 여러 장, 효과 없이 바로 넘어감"""
    num = str(r.get("no", "")).strip()
    kind = (r.get("kind") or "일반").strip() or "일반"
    raw = st.get("_raw") or {}
    mode = st.get("fc_num") or "구분별 색"
    ncol = FC_COLORS[FC_KIND.get(kind, "하늘")] if mode in ("구분별 색", "(기본)") else FC_COLORS.get(mode, FC_COLORS["하늘"])
    if raw.get("num_big") not in (None, "", "(기본)") and COLORS.get(raw["num_big"]):
        ncol = COLORS[raw["num_big"]]
    dmode = st.get("fc_dir") or "초록"
    dcol = ncol if dmode == "번호색과 같게" else FC_COLORS.get(dmode, FC_COLORS["초록"])
    white = FC_COLORS["흰색"]
    tcol = COLORS.get(raw.get("end")) if raw.get("end") not in (None, "", "(기본)") and COLORS.get(raw.get("end")) else white
    sep_white = (st.get("fc_sep") or "초록 –") == "흰색 ▸"
    br = r.get("_br") or {}
    pages = []

    def new():
        return Image.new("RGBA", (OLD_W, OLD_H), (0, 0, 0, 255))

    def add(img, k, ms=None):
        im = img.convert("RGB")
        im.info["duration"] = ms or OLD_MS.get(k, 1500)
        im.info["page_kind"] = k
        im.info["trans"] = "cut"
        pages.append(im)

    # 로고(있으면 번호 왼쪽) — 사진의 로고 자리. 이미지는 사용자가 가진 파일을 씀
    logo_im = None
    logo = (r.get("logo") or "").strip()
    if logo:
        lp = Path(logo).expanduser()
        if lp.exists():
            src = Image.open(lp).convert("RGBA")
            src.thumbnail((40, 30), Image.LANCZOS)
            logo_im = logo_to_led(lp, src.width, src.height, full=True)
        else:
            WARN.append((f"로고 파일 없음: {logo}", 0))
    lw = (logo_im.width + 3) if logo_im else 0
    # 사진처럼 판 높이를 채우는 좁고 긴 숫자 (148-1이 판 폭의 40% 남짓)
    n_img = render_num_orig(num, ncol, 30, 76 - lw, kx=0.78, th=100, track=1, warn=False) if num else None
    x0 = lw + (n_img.width + 5 if n_img else 2)
    room = OLD_W - 1 - x0

    def head():
        img = new()
        if logo_im:
            img.alpha_composite(logo_im, (0, (OLD_H - logo_im.height) // 2))
        if n_img:
            img.alpha_composite(n_img, (lw, (OLD_H - n_img.height) // 2))
        return img

    def dir_label(w):
        ims = [_old_text(ch, dcol, "two", 10 ** 4) for ch in "(방향)"]
        tot = sum(i.width for i in ims)
        gap = max(2, min(14, (int(w * 0.8) - tot) // 3))     # 사진처럼 칸 폭만큼 벌림: ( 방 향 )
        lab = Image.new("RGBA", (tot + gap * 3, max(i.height for i in ims)), (0, 0, 0, 0))
        xx = 0
        for i in ims:
            lab.alpha_composite(i, (xx, lab.height - i.height))
            xx += i.width + gap
        return lab if lab.width <= w else squeeze_w(lab, w)

    end_ko = (r.get("end_ko") or "").strip()
    start_ko = (r.get("start_ko") or "").strip()
    if end_ko:                                         # ① 번호 + 행선지 + (방 향)
        img = head()
        one = _old_text(end_ko, tcol, "top", 10 ** 4)
        if len(end_ko) <= 5 and one.width <= room:
            t = _spaced(end_ko, tcol, "top", room)
        else:
            t = _fc_text_fit(end_ko, tcol, ("top", "two", "via"), room)
        img.alpha_composite(t, (x0 + (room - t.width) // 2, 1))
        lab = dir_label(room)
        img.alpha_composite(lab, (x0 + (room - lab.width) // 2, 18))
        add(img, "route", 2000)

    def route_page(a, b, fa=None, fb=None):
        img = head()
        sw = 10
        def nat(t):
            return _fc_text_fit(t, white, ("two",), 10 ** 4).width if t else 1
        wa, wb = nat(a), nat(b)                        # 글자 길이에 맞춰 좌우 칸을 나눔 (30~70%)
        half = int((room - sw) * min(0.7, max(0.3, wa / (wa + wb))))
        _fc_block(a, white, x0, x0 + half - 1, img, fa)
        mx = x0 + half + 1
        d = ImageDraw.Draw(img)
        if sep_white:
            d.polygon([(mx + 2, 12), (mx + 7, 16), (mx + 2, 20)], fill=white + (255,))
        else:
            d.rectangle([mx, 15, mx + 7, 16], fill=FC_COLORS["초록"] + (255,))
        _fc_block(b, white, x0 + half + sw + 1, OLD_W - 1, img, fb)
        add(img, "route", 1800)

    if start_ko and end_ko:                            # ② 번호 + 기점 – 종점
        route_page(start_ko, end_ko, br.get("start_ko"), br.get("end_ko"))
    s_en = (r.get("start_en") or "").strip() or (romanize_name(start_ko) if start_ko else "")
    e_en = (r.get("end_en") or "").strip() or (romanize_name(end_ko) if end_ko else "")
    if s_en and e_en:                                  # ③ 영문
        route_page(s_en, e_en, br.get("start_en"), br.get("end_en"))
    elif e_en:
        img = head()
        t = _fc_text_fit(e_en, tcol, ("top", "two"), room)
        img.alpha_composite(t, (x0 + (room - t.width) // 2, 1))
        lab = _fc_text_fit("(Direction)", dcol, ("two", "via"), room)
        img.alpha_composite(lab, (x0 + (room - lab.width) // 2, 18))
        add(img, "route", 1800)
    if num:                                            # ④ 번호만 크게 (사진 '5 0')
        img = new()
        n = render_num_orig(num, ncol, 30, OLD_W - 8, kx=1.0, th=100, track=1, warn=False)
        if len(num) == 2 and num.isdigit():             # 두 자리 번호는 사진('5 0')처럼 글자 사이를 넓게
            digs = [render_num_orig(ch, ncol, 30, 10 ** 4, kx=1.0, th=100, track=0, warn=False) for ch in num]
            tot = sum(d_.width for d_ in digs)
            gap = max(4, min(36, OLD_W - 60 - tot))
            n = Image.new("RGBA", (tot + gap * (len(digs) - 1), max(d_.height for d_ in digs)), (0, 0, 0, 0))
            xx = 0
            for d_ in digs:
                n.alpha_composite(d_, (xx, 0))
                xx += d_.width + gap
        img.alpha_composite(n, ((OLD_W - n.width) // 2, (OLD_H - n.height) // 2))
        add(img, "big", 1500)
    seats = seats_text(r)
    if seats:
        add(seats_page(seats, white, ncol), "seats")
    if r.get("slogan", "").strip():
        img = new()
        t = _fc_text_fit(r["slogan"].strip(), white, ("big", "top", "two"), OLD_W - 8)
        img.alpha_composite(t, ((OLD_W - t.width) // 2, (OLD_H - t.height) // 2))
        add(img, "slogan", 2000)
    return pages


TEMPLATES = {
    "jinmyeong_front_192x160": tpl_jinmyeong_front,
    "jinmyeong_uniform_192x160": tpl_jinmyeong_uniform,
    "simple_front_192x160": tpl_simple_front,
    "busan_old_192x32 (부산 구형 3색)": tpl_busan_old,
    "busan_new_192x32 (부산 신형 풀컬러)": tpl_busan_new,
    "fullcolor_dir_192x32 (풀컬러 방향)": tpl_fullcolor_dir,
}


LINE_BREAK_KEYS = ("start_ko", "end_ko", "start_en", "end_en", "start_ja", "end_ja", "start_zh", "end_zh")


def split_breaks(route):
    """기점·종점 칸의 '|'(또는 줄바꿈)는 '두 줄일 때 여기서 나눔' 표시.
    한 줄로 쓰는 곳에는 표시를 지운 글자를, 두 줄 화면에는 나눌 위치를 따로 넘김."""
    r = dict(route)
    br = dict(r.get("_br") or {})               # 이미 한 번 나눈 노선이면 그 정보를 유지
    for k in LINE_BREAK_KEYS:
        v = str(r.get(k, "") or "")
        if "|" in v or "\n" in v:
            parts = [p.strip() for p in v.replace("\n", "|").split("|") if p.strip()]
            if len(parts) >= 2:
                br[k] = (parts[0], " ".join(parts[1:]) if k.endswith(("_en",)) else "".join(parts[1:]))
            sep = " " if k.endswith("_en") else ""
            r[k] = sep.join(parts)
    r["_br"] = br
    return r


def render_route(route, style=None):
    route = split_breaks(route)
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
                if " " not in text:
                    small = render_mono([(text, color)], h2, f2, s2, 10 ** 4)
                    if small.width <= bw or not has_hangul(text):
                        li = render_mono([(text, color)], h2, f2, s2, bw)
                        img.alpha_composite(li, (x + (bw - li.width) // 2, y + (bh - h2) // 2))
                        continue
                a, b2 = split_two(text)
                gap = max(1, (bh - 2 * h2) // 3)
                for i, part in enumerate((a, b2)):
                    li = render_mono([(part, color)], h2, f2, s2, bw)
                    img.alpha_composite(li, (x + (bw - li.width) // 2, y + gap + i * (h2 + gap)))
    return img.convert("RGB")


def render_pages(route, style=None, with_edits=True):
    """노선 → [페이지 이미지...]. 점 편집까지 적용. WARN 에 압축 경고가 쌓임."""
    WARN.clear()
    route = split_breaks(route)                   # '김해|구산동' 표시는 JSON 판에서도 지움
    tname = route.get("template") or ""
    jt = load_json_templates().get(tname)
    if jt is None:
        res = render_route(route, style)
        pages = res if isinstance(res, list) else [res]
        if not pages:                             # 내용이 하나도 없는 노선: 빈 화면 1장 (미리보기·저장이 깨지지 않게)
            pages = [Image.new("RGB", (OLD_W, OLD_H))]
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


# ───────── 행선지 중국어·일본어 (신형 판 '다대포행' 밑줄) ─────────
# 1) 노선의 종점(중)·종점(일) 칸  2) ~/BusSignMaker/translate_dict.json  3) 아래 기본 사전
# 4) 일본어는 그래도 없으면 가타카나로 소리 나는 대로 (다대포 → タデポ). 중국어는 한자를 알 수 없으면 생략.
TRANS_FILE = APP_DIR / "translate_dict.json"
TRANS_DEFAULT = {
    "부산": ("釜山", "釜山"), "다대포": ("多大浦", "多大浦"), "다대포해수욕장": ("多大浦海水浴场", "多大浦海水浴場"),
    "해운대": ("海云台", "海雲台"), "해운대해수욕장": ("海云台海水浴场", "海雲台海水浴場"),
    "광안리": ("广安里", "広安里"), "광안리해수욕장": ("广安里海水浴场", "広安里海水浴場"),
    "송정": ("松亭", "松亭"), "기장": ("机张", "機張"), "서면": ("西面", "西面"), "남포동": ("南浦洞", "南浦洞"),
    "부산역": ("釜山站", "釜山駅"), "부산진시장": ("釜山镇市场", "釜山鎮市場"), "자갈치": ("札嘎其", "チャガルチ"),
    "자갈치시장": ("札嘎其市场", "チャガルチ市場"), "사상": ("沙上", "沙上"), "하단": ("下端", "下端"),
    "동래": ("东莱", "東莱"), "온천장": ("温泉场", "温泉場"), "노포동": ("老圃洞", "老圃洞"), "범어사": ("梵鱼寺", "梵魚寺"),
    "센텀시티": ("Centum City", "センタムシティ"), "벡스코": ("BEXCO", "BEXCO"), "김해공항": ("金海机场", "金海空港"),
    "김해": ("金海", "金海"), "구포": ("龟浦", "亀浦"), "덕천": ("德川", "徳川"), "명지": ("鸣旨", "鳴旨"),
    "녹산": ("菉山", "菉山"), "영도": ("影岛", "影島"), "태종대": ("太宗台", "太宗台"), "감천문화마을": ("甘川文化村", "甘川文化村"),
    "신라대": ("新罗大学", "新羅大学"), "부산대": ("釜山大学", "釜山大学"), "동아대": ("东亚大学", "東亜大学"),
    "부경대": ("釜庆大学", "釜慶大学"), "경성대": ("庆星大学", "慶星大学"), "연산": ("莲山", "蓮山"),
    "수영": ("水营", "水営"), "기장시장": ("机张市场", "機張市場"), "반여동": ("盘如洞", "盤如洞"), "반여4동": ("盘如4洞", "盤如4洞"),
    "사하구청": ("沙下区厅", "沙下区庁"), "부산시청": ("釜山市厅", "釜山市庁"), "용호동": ("龙湖洞", "龍湖洞"),
    # 부산 동·지명 추가 (기본 사전은 참고용: 정확한 표기는 칸에 직접 적거나 translate_dict.json에서 고치세요)
    "범일동": ("凡一洞", "凡一洞"), "구산동": ("龟山洞", "亀山洞"), "좌천동": ("佐川洞", "佐川洞"), "초량": ("草梁", "草梁"), "수정동": ("水晶洞", "水晶洞"), "대연동":
    ("大渊洞", "大淵洞"), "광안동": ("广安洞", "広安洞"), "민락동": ("民乐洞", "民楽洞"), "재송동": ("栽松洞", "栽松洞"), "반송": ("盘松", "盤松"), "정관": ("鼎冠",
    "鼎冠"), "장안": ("长安", "長安"), "구서동": ("久瑞洞", "久瑞洞"), "장전동": ("长箭洞", "長箭洞"), "온천동": ("温泉洞", "温泉洞"), "사직동": ("社稷洞", "社稷洞"),
    "거제동": ("巨堤洞", "巨堤洞"), "연산동": ("莲山洞", "蓮山洞"), "양정": ("杨亭", "楊亭"), "부전동": ("釜田洞", "釜田洞"), "전포동": ("田浦洞", "田浦洞"), "당감동":
    ("堂甘洞", "堂甘洞"), "개금": ("开琴", "開琴"), "주례": ("周礼", "周礼"), "감전": ("甘田", "甘田"), "학장": ("鹤章", "鶴章"), "엄궁": ("严弓", "厳弓"),
    "괴정": ("槐亭", "槐亭"), "당리": ("堂里", "堂里"), "신평": ("新平", "新平"), "장림": ("长林", "長林"), "다대동": ("多大洞", "多大洞"), "감천": ("甘川",
    "甘川"), "송도": ("松岛", "松島"), "남항": ("南港", "南港"), "청학동": ("青鹤洞", "青鶴洞"), "동삼동": ("东三洞", "東三洞"), "금곡": ("金谷", "金谷"), "화명":
    ("华明", "華明"), "만덕": ("万德", "万徳"), "강서구청": ("江西区厅", "江西区庁"), "금정구청": ("金井区厅", "金井区庁"), "해운대구청": ("海云台区厅", "海雲台区庁"),
    "중앙동": ("中央洞", "中央洞"), "충무동": ("忠武洞", "忠武洞"), "대신동": ("大新洞", "大新洞"), "서대신동": ("西大新洞", "西大新洞"), "동대신동": ("东大新洞",
    "東大新洞"), "아미동": ("峨嵋洞", "峨嵋洞"), "영주동": ("瀛州洞", "瀛州洞"), "문현동": ("门岘洞", "門峴洞"), "용당동": ("龙塘洞", "龍塘洞"), "우동": ("佑洞",
    "佑洞"), "중동": ("中洞", "中洞"), "송정동": ("松亭洞", "松亭洞"), "기장군청": ("机张郡厅", "機張郡庁"), "일광": ("日光", "日光"), "월내": ("月内", "月内"),
}
_TRANS_SUFFIX = {"역": ("站", "駅"), "시장": ("市场", "市場"), "대학교": ("大学", "大学"), "해수욕장": ("海水浴场", "海水浴場"),
                 "공영차고지": ("公营车库", "公営車庫"), "차고지": ("车库", "車庫"), "터미널": ("客运站", "ターミナル"),
                 "아파트": ("公寓", "アパート"), "시청": ("市厅", "市庁"), "구청": ("区厅", "区庁")}


def load_trans_dict():
    d = {k: {"zh": v[0], "ja": v[1]} for k, v in TRANS_DEFAULT.items()}
    try:
        for k, v in json.loads(TRANS_FILE.read_text(encoding="utf-8")).items():
            if isinstance(v, dict) and not k.startswith("_"):
                d[k] = {**d.get(k, {}), **v}
    except Exception:
        pass
    return d


_V = ["a", "e", "ya", "ye", "o", "e", "yo", "ye", "o", "wa", "we", "we", "yo", "u", "wo", "we", "wi", "yu", "u", "ui", "i"]
_K5 = {"": "アイウエオ", "k": "カキクケコ", "g": "ガギグゲゴ", "s": "サシスセソ", "t": "タチトゥテト", "d": "ダジドゥデド",
       "n": "ナニヌネノ", "h": "ハヒフヘホ", "b": "バビブベボ", "p": "パピプペポ", "m": "マミムメモ", "r": "ラリルレロ",
       "ch": "チャチチュチェチョ", "j": "ジャジジュジェジョ"}


def _kana(c, v):
    def row(cc, i):
        r_ = _K5[cc]
        if cc in ("ch", "j"):
            return [r_[0:2], r_[2], r_[3:5], r_[5:7], r_[7:9]][i]
        if cc in ("t", "d"):
            return [r_[0], ("ティ" if cc == "t" else "ディ"), r_[2:4], r_[4], r_[5]][i]
        return r_[i]
    idx = {"a": 0, "i": 1, "u": 2, "e": 3, "o": 4}
    if v in idx:
        return row(c, idx[v])
    if v.startswith("y"):                    # 야·유·요·예
        base = {"ya": "ャ", "yu": "ュ", "yo": "ョ", "ye": "ェ"}[v]
        if c in ("", ):
            return {"ya": "ヤ", "yu": "ユ", "yo": "ヨ", "ye": "イェ"}[v]
        if c in ("ch", "j"):
            return row(c, {"ya": 0, "yu": 2, "yo": 4, "ye": 3}[v])
        return row(c, 1) + base
    if v.startswith("w"):                    # 와·워·웨·위
        if c == "":
            return {"wa": "ワ", "wo": "ウォ", "we": "ウェ", "wi": "ウィ"}[v]
        return row(c, 2) + {"wa": "ァ", "wo": "ォ", "we": "ェ", "wi": "ィ"}[v]
    if v == "ui":
        return row(c, 2) + "イ"
    return row(c, 0)


def hangul_to_katakana(text):
    """한글 → 가타카나 (소리 나는 대로, 첫소리 ㄱㄷㅂㅈ은 맑은소리: 부산 → プサン)"""
    on = ["k", "k", "n", "t", "t", "r", "m", "p", "p", "s", "s", "", "ch", "ch", "ch", "k", "t", "p", "h"]
    voiced = {"k": "g", "t": "d", "p": "b", "ch": "j"}
    tail = {1: "ク", 2: "ク", 3: "ク", 4: "ン", 5: "ン", 6: "ン", 7: "ッ", 8: "ル", 9: "ル", 10: "ル", 11: "ル", 12: "ル",
            13: "ル", 14: "ル", 15: "ル", 16: "ム", 17: "プ", 18: "プ", 19: "ッ", 20: "ッ", 21: "ン", 22: "ッ", 23: "ッ",
            24: "ク", 25: "ッ", 26: "プ", 27: "ッ"}
    out, prev_vowel = "", False
    for ch in text:
        o = ord(ch) - 0xAC00
        if not 0 <= o < 11172:
            out += ch
            prev_vowel = False
            continue
        L, V, T = o // 588, (o % 588) // 28, o % 28
        c = on[L]
        if prev_vowel and L in (0, 3, 7, 12):        # 가운데 오는 ㄱㄷㅂㅈ은 흐린소리 (해운대 → ヘウンデ)
            c = voiced[c]
        out += _kana(c, _V[V])
        if T:
            out += tail.get(T, "")
        prev_vowel = T in (0, 4, 8, 16, 21)          # 모음·ㄴ·ㄹ·ㅁ·ㅇ 뒤
    return out


def translate_place(name, lang):
    """행선지 → 중국어(lang='zh')·일본어('ja'). 모르면 일본어는 가타카나, 중국어는 ''"""
    name = (name or "").strip()
    if not name:
        return ""
    d = load_trans_dict()
    if name in d and d[name].get(lang):
        return d[name][lang]
    li = 0 if lang == "zh" else 1
    for suf in sorted(_TRANS_SUFFIX, key=len, reverse=True):
        if name.endswith(suf) and len(name) > len(suf):
            stem = name[: -len(suf)]
            if stem in d and d[stem].get(lang):
                return d[stem][lang] + _TRANS_SUFFIX[suf][li]
            if lang == "ja":
                return hangul_to_katakana(stem) + _TRANS_SUFFIX[suf][1]
            return ""
    return hangul_to_katakana(name) if lang == "ja" else ""


ROMAN_FILE = APP_DIR / "romanize_dict.json"
ROMAN_DEFAULT = {
    "_설명": "exact = 이름 그대로 바꿈, suffix = 끝말 바꿈 (긴 것부터 적용). 자유롭게 고치세요.",
    "exact": {"서울역": "Seoul Station", "부산역": "Busan Station", "서면": "Seomyeon"},
    "suffix": {"관광단지": " Tourist Complex", "공영차고지": " Bus Garage", "차고지": " Garage",
               "산업단지": " Industrial Complex", "산단": " Industrial Complex", "의학원": " Medical Center",
               "환승센터": " Transfer Center", "버스터미널": " Bus Terminal", "터미널": " Terminal",
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
        if len(word) >= 2 and word[-1] in "동리읍면" and re.fullmatch(r"[가-힣]+", word):
            # 행정구역 단위는 붙임표 (국어의 로마자 표기법 제5항): 구산동 → Gusan-dong
            parts.append(_cap(romanize_word(word[:-1])) + "-" + romanize_word(word[-1]))
            continue
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
    wd = len(str(n_pages))          # 10장 이상이면 01, 02… (정렬 순서가 맞도록)
    for i in range(n_pages):
        vals["page"] = str(i + 1).zfill(wd)
        base = pattern.format_map(vals)
        if n_pages > 1 and "{page}" not in pattern:
            base += f"_{str(i + 1).zfill(wd)}"
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


# ───────────────────────── LED 모양 미리보기 · 움직이는 GIF ─────────────────────────
def led_look(img, z=6, glow=True):
    """점 하나를 둥근 LED로 그림. 꺼진 점은 어둡게 보여서 실제 전광판처럼 보임"""
    import numpy as np
    a = np.asarray(img.convert("RGB")).astype(np.float32)
    h, w, _ = a.shape
    yy, xx = np.mgrid[0:z, 0:z]
    r = (xx - (z - 1) / 2) ** 2 + (yy - (z - 1) / 2) ** 2
    dot = (r <= (z * 0.42) ** 2).astype(np.float32)
    big = np.repeat(np.repeat(a, z, 0), z, 1)
    lit = (big.max(2, keepdims=True) > 30).astype(np.float32)
    mask = np.tile(dot, (h, w))[..., None]
    off = np.array([38, 30, 28], np.float32)
    out = np.where(lit > 0, big, off) * mask + (1 - mask) * np.array([12, 10, 10], np.float32)
    im = Image.fromarray(np.clip(out, 0, 255).astype(np.uint8), "RGB")
    if glow:
        from PIL import ImageFilter
        halo = Image.fromarray(np.clip(big * lit, 0, 255).astype(np.uint8), "RGB").filter(ImageFilter.GaussianBlur(z * 0.7))
        im = Image.blend(im, Image.fromarray(np.maximum(np.asarray(im), np.asarray(halo) // 2).astype(np.uint8)), 0.6)
    return im


def page_duration(img, default=1000):
    return int(img.info.get("duration", default))


def anim_frames(pages, scroll_steps=6, band=18, step_ms=50):
    """페이지 → 재생 프레임 [(이미지, ms)].
    부산 구형·신형 모두 넘김 효과 없이(PPT 전환 없음처럼) 바로 다음 화면으로 바뀜."""
    return [(p, page_duration(p)) for p in pages]


def gif_frames(pages, z=4, led=True, scroll_steps=6):
    out = []
    for f, ms in anim_frames(pages, scroll_steps):
        out.append((led_look(f, z) if led else f.resize((f.width * z, f.height * z), Image.NEAREST), ms))
    return out


def save_gif(pages, path, z=4, led=True):
    fr = gif_frames(pages, z, led)
    if not fr:
        return 0
    ims = [f.convert("P", palette=Image.Palette.ADAPTIVE, colors=128) for f, _ in fr]
    ims[0].save(path, save_all=True, append_images=ims[1:], duration=[ms for _, ms in fr], loop=0, optimize=False, disposal=1)
    return len(ims)

# ───────────────────────── 엑셀 ─────────────────────────
XL_COLS = [("no", "번호"), ("start_ko", "기점(한)"), ("end_ko", "종점(한)"),
           ("start_en", "기점(영)"), ("end_en", "종점(영)"), ("via", "경유지"),
           ("template", "템플릿"), ("kind", "구분"),
           ("start_ja", "기점(일)"), ("end_ja", "종점(일)"), ("start_zh", "기점(중)"), ("end_zh", "종점(중)"),
           ("slogan", "문구"), ("logo", "로고 이미지"), ("seats", "잔여좌석")]


SAMPLE_ROUTES = [
    # (번호, 기점한, 종점한, 기점영, 종점영, 경유지, 템플릿, 구분, 설명)
    ("360", "서울역 환승센터", "강일IC", "Seoul Station", "Gang-il IC",
     "동묘앞,을지로4가\n강동역,십자성마을\nDongmyo,Euljiro 4-ga\nGangdong Stn,Sipjaseong Village",
     "jinmyeong_front_192x160", "일반", "기본 예시 (원본 화면과 같은 노선)"),
    ("1100", "부산역", "하단역", "Busan Station", "Hadan Stn",
     "남포동,자갈치\nNampodong,Jagalchi", "jinmyeong_front_192x160", "일반", "경유지 한 줄씩"),
    ("1003", "사하구청", "해운대해수욕장", "", "",
     "괴정역,대티역,서면\n부전역,연산역", "jinmyeong_front_192x160", "급행",
     "영문 칸이 비어 있음 → [모든 노선 채우기]로 자동 변환해 보세요. 급행은 번호가 빨간 상자로 반전"),
    ("58-1", "괴정", "다대포해수욕장", "Goejeong", "Dadaepo Beach",
     "", "jinmyeong_front_192x160", "좌석", "경유지 없음, 좌석은 번호가 빨강"),
    ("심야1", "부산역", "해운대", "Busan Station", "Haeundae",
     "서면,연산역", "jinmyeong_front_192x160", "심야", "심야는 하늘색 반전 · 번호에 한글"),
    ("사하10", "괴정역", "감천문화마을", "Goejeong Stn", "Gamcheon Culture Village",
     "", "side_160x32 (측면, 한/영 2장)", "마을", "측면판: 한글/영문 2장으로 저장 (_1, _2)"),
    ("86", "기장", "부산대학교", "Gijang", "Pusan National Univ.",
     "송정,반여농산물시장,동래역,부산대역\nSongjeong,Banyeo Mkt,Dongnae,PNU Stn",
     "jinmyeong_front_192x160", "일반", "일부러 긴 경유지 → 압축 경고가 뜨는 예"),
    ("1004", "", "", "", "", "", "rear_64x32 (후면 번호)", "일반", "후면 번호판 64x32"),
    ("1001", "신평", "송정", "Sinpyeong", "Songjeong",
     "", "front_192x32 (전면 한 줄, 한/영 2장)", "순환", "전면 한 줄판 192x32, 순환은 주황"),
    ("37", "금정공영차고지", "고원아파트", "Geumjeong Bus Garage", "Gowon APT.",
     "노포동터미널,월평삼거리,정관,좌천,임랑삼거리\nNopo-dong Bus Term.,Wolpyeong Samgeori,Jeonggwan,Jwacheon,Imnang Samgeori",
     "busan_old_192x32 (부산 구형 3색)", "일반",
     "부산 구형 3색 LED 192x32: 큰 번호 → 경유지(한/영 교대) → 행선(한/영/일/중). 미리보기 [▶ 재생]",
     {"start_ja": "金井コン ヨン・チャゴジ", "end_ja": "Gowonアパート", "start_zh": "金井 公營車庫", "end_zh": "Gowon公寓"}),
    ("187", "반송", "대룡마을", "Bansong", "Daeryong Village",
     "원자력의학원,장안산단,명례산단\nInstitute of radiological & medical sciences,Jang-an Industrial Complex,Myeongnye Industrial Complex",
     "busan_old_192x32 (부산 구형 3색)", "일반",
     "긴 영문 경유지는 두 화면으로 나눠 읽기 쉽게. '로고 이미지' 칸에 가진 이미지 파일을 넣으면 한 장면으로 들어감", {}),
    ("99", "남포동", "부산진시장", "Nampo-dong", "Busanjin Market",
     "충무교차로,시청,서면,범내골\nChungmu Intersection,City Hall,Seomyeon,Beomnaegol",
     "busan_new_192x32 (부산 신형 풀컬러)", "일반",
     "부산 신형 풀컬러 LED: 큰 번호 → 번호 + 'OO행' → 'O O 행' + 영문 → 작은 경유지 → 문구(분홍)",
     {"slogan": "Busan is good"}),
    ("2029", "녹산", "해운대", "Noksan", "Haeundae",
     "명지,하단,괴정,서면,센텀시티\nMyeongji,Hadan,Goejeong,Seomyeon,Centum City",
     "busan_new_192x32 (부산 신형 풀컬러)", "급행",
     "신형: 흰 번호 → 번호+'OO행' → 'O O 행'+영문·중문 → 효과 없이 바로 바뀜. '잔여좌석' 칸에 숫자를 넣으면 잔여좌석 화면 추가",
     {"seats": "24", "slogan": "Busan is good"}),
    ("2000", "녹산", "해운대", "Noksan", "Haeundae",
     "명지,하단,서면\nMyeongji,Hadan,Seomyeon",
     "busan_old_192x32 (부산 구형 3색)", "일반",
     "구형 3색에도 잔여좌석 화면(빨간 '잔여좌석' + 주황 숫자)을 넣을 수 있음",
     {"seats": "17"}),
    ("148-1", "학장동", "부산대학교", "", "",
     "", "fullcolor_dir_192x32 (풀컬러 방향)", "일반",
     "풀컬러 방향 표시: 번호(구분별 색) + 행선지 + 초록 '(방 향)' → 번호 + 기점 – 종점 → 영문 → 큰 번호. 스타일에서 번호색·(방향)색·구분자(– / ▸) 선택",
     {}),
    ("864", "인경리", "청주의료원", "Ingyeong-ri", "Cheongju Medical Center",
     "", "fullcolor_dir_192x32 (풀컬러 방향)", "좌석",
     "긴 행선지는 사진처럼 두 줄로 ('청주 / 의료원'). '|'로 줄 바꿀 곳을 직접 정할 수도 있음",
     {}),
]


def make_sample_excel(path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "노선"
    ws.append([c[1] for c in XL_COLS] + ["설명(읽지 않음)"])
    base_keys = ["no", "start_ko", "end_ko", "start_en", "end_en", "via", "template", "kind"]
    for row in SAMPLE_ROUTES:
        d = dict(zip(base_keys, row[:8]))
        d.update(row[9] if len(row) > 9 else {})
        ws.append([d.get(k, "") for k, _ in XL_COLS] + [row[8]])
    head = PatternFill("solid", fgColor="FFE699")
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = head
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    from openpyxl.utils import get_column_letter
    widths = [8, 16, 16, 16, 22, 46, 30, 7, 14, 14, 14, 14, 12, 14, 9, 44]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    g = wb.create_sheet("사용법")
    for line in [
        ["노선이미지 생성기 엑셀 형식"],
        [""],
        ["· 첫 번째 시트(노선)의 첫 줄은 제목 줄입니다. 열 순서는 바뀌어도 되고, 제목 글자만 맞으면 됩니다."],
        ["· 번호: 행선판에 나올 노선 번호 (숫자·하이픈은 원본형 번호 글꼴, 한글이 섞이면 일반 글꼴)"],
        ["· 기점/종점(한·영): 공백이 있으면 두 줄로 나뉘고, 짧으면 한 줄로 크게 나옵니다."],
        ["· 경유지: 쉼표 = 정류장, 줄바꿈(Alt+Enter / 맥은 ⌃⌥Enter) = 줄. 한글 줄 다음에 영문 줄."],
        ["· 템플릿: jinmyeong_front_192x160 / busan_old_192x32 (부산 구형 3색) / busan_new_192x32 (부산 신형 풀컬러) / fullcolor_dir_192x32 (풀컬러 방향) / side_160x32 (측면, 한/영 2장) / front_192x32 (전면 한 줄, 한/영 2장) / rear_64x32 (후면 번호)"],
        ["· 구분: 일반 / 급행 / 좌석 / 심야 / 마을 / 순환 (색·반전은 프로그램의 스타일 탭에서 바꿈)"],
        ["· 기점/종점(일·중), 문구, 로고 이미지: 부산 구형 3색 템플릿(busan_old_192x32)의 행선 화면·문구 화면·로고 화면에 쓰임 (비워 두면 그 화면은 생략)"],
        ["· 로고 이미지: PNG·BMP·JPG 파일 경로. 판 크기에 맞춰 줄여 넣습니다(구형은 빨강·초록·주황 3색으로 바꿈)."],
        ["· 잔여좌석: 숫자(예: 24)를 넣으면 부산 구형·신형 템플릿에 '잔여좌석 24' 화면이 추가됩니다. 비워 두면 생략."],
        ["· 영문 칸을 비워 두고 불러온 뒤 [편집 탭 → 모든 노선 채우기]를 누르면 자동으로 채워집니다."],
        ["· '설명' 열처럼 제목이 다른 열은 무시됩니다."],
    ]:
        g.append(line)
    g.column_dimensions["A"].width = 120
    g["A1"].font = Font(bold=True, size=14)
    wb.active = 0
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
    from openpyxl.utils import get_column_letter
    for i, w in enumerate([8, 18, 14, 16, 14, 50, 26, 8, 14, 14, 14, 14, 12, 14, 9], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
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
EXTRA_FIELDS = [("start_ja", "기점(日)"), ("end_ja", "종점(日)"), ("start_zh", "기점(中)"), ("end_zh", "종점(中)")]
GLYPH_SIZES_TEXT = {"14 · 두 줄 기점/종점": (14, 13, 7), "24 · 한 줄 한글 종점": (24, 23, 12),
                    "13 · 영문 한 줄 / 경유지": (13, 13, 7), "11 · 작은 글자": (11, 11, 6), "9 · 아주 작은 글자": (9, 9, 5)}
GLYPH_SIZES_NUM = {"22 · 큰 번호": 22, "19 · 작은 번호": 19, "24 · 후면 번호": 24}
IS_MAC = sys.platform == "darwin"
MOD = "Command" if IS_MAC else "Control"


# ───────────────────────── 화면 테마 (다크 / 라이트) ─────────────────────────
THEMES = {
    "다크": {"bg": "#0f1115", "surface": "#171a21", "surface2": "#20242d", "border": "#2b303b",
             "text": "#e7e9ee", "muted": "#8a92a6", "accent": "#4f8cff", "accent_hi": "#6aa0ff",
             "on_accent": "#ffffff", "warn": "#ff6b6b", "select": "#2b4a85", "header": "#12151b",
             "canvas": "#08090b", "input": "#11141a"},
    "라이트": {"bg": "#f3f4f7", "surface": "#ffffff", "surface2": "#eef0f4", "border": "#dde1e8",
              "text": "#1c2230", "muted": "#687186", "accent": "#2563eb", "accent_hi": "#3b76f0",
              "on_accent": "#ffffff", "warn": "#d93636", "select": "#d6e4ff", "header": "#ffffff",
              "canvas": "#14161a", "input": "#ffffff"},
}


def setup_theme(root, name, ttk):
    """ttk 'clam'을 바탕으로 색·여백·글꼴을 요즘 앱처럼 맞춤. 몇 번이든 다시 불러도 됨(테마 전환)"""
    import tkinter.font as tkfont
    P = THEMES.get(name) or THEMES["다크"]
    st = ttk.Style(root)
    st.theme_use("clam")
    base = tkfont.nametofont("TkDefaultFont")
    fam = base.actual("family")
    size = 13 if IS_MAC else 10
    for fn in ("TkDefaultFont", "TkTextFont", "TkMenuFont", "TkHeadingFont"):
        try:
            tkfont.nametofont(fn).configure(size=size)
        except Exception:
            pass
    F = {"base": (fam, size), "bold": (fam, size, "bold"), "title": (fam, size + 3, "bold"),
         "section": (fam, size, "bold"), "small": (fam, size - 1), "list": (fam, size + 1)}
    st.configure(".", background=P["bg"], foreground=P["text"], bordercolor=P["border"],
                 lightcolor=P["surface"], darkcolor=P["surface"], troughcolor=P["surface2"],
                 fieldbackground=P["input"], selectbackground=P["select"],
                 selectforeground=P["text"], insertcolor=P["text"], arrowcolor=P["muted"],
                 focuscolor=P["accent"], font=F["base"])
    for sty, bg in (("TFrame", P["surface"]), ("Card.TFrame", P["surface"]), ("Header.TFrame", P["header"]),
                    ("Status.TFrame", P["header"])):
        st.configure(sty, background=bg)
    st.configure("Card.TFrame", bordercolor=P["border"], relief="solid", borderwidth=1)
    st.configure("TLabel", background=P["surface"], foreground=P["text"])
    st.configure("Bg.TLabel", background=P["bg"])
    st.configure("Header.TLabel", background=P["header"], foreground=P["text"])
    st.configure("Brand.TLabel", background=P["header"], foreground=P["text"], font=F["title"])
    st.configure("BrandSub.TLabel", background=P["header"], foreground=P["muted"], font=F["small"])
    st.configure("Status.TLabel", background=P["header"], foreground=P["muted"], font=F["small"], padding=(12, 5))
    st.configure("Muted.TLabel", background=P["surface"], foreground=P["muted"], font=F["small"])
    st.configure("Warn.TLabel", background=P["surface"], foreground=P["warn"], font=F["small"])
    st.configure("Section.TLabel", background=P["surface"], foreground=P["text"], font=F["section"])
    st.configure("Badge.TLabel", background=P["surface2"], foreground=P["muted"], font=F["small"], padding=(8, 1))
    # 카드 안 묶음 상자
    st.configure("TLabelframe", background=P["surface"], bordercolor=P["border"], relief="solid",
                 borderwidth=1, lightcolor=P["border"], darkcolor=P["border"])
    st.configure("TLabelframe.Label", background=P["surface"], foreground=P["accent"], font=F["bold"])
    # 버튼
    st.configure("TButton", background=P["surface2"], foreground=P["text"], bordercolor=P["border"],
                 lightcolor=P["surface2"], darkcolor=P["surface2"], padding=(10, 5), relief="flat",
                 focusthickness=0, focuscolor=P["surface2"], width=-4)
    st.map("TButton", background=[("pressed", P["border"]), ("active", P["border"])],
           bordercolor=[("active", P["accent"])])
    st.configure("Accent.TButton", background=P["accent"], foreground=P["on_accent"], bordercolor=P["accent"],
                 lightcolor=P["accent"], darkcolor=P["accent"], font=F["bold"], focuscolor=P["accent"])
    st.map("Accent.TButton", background=[("pressed", P["accent"]), ("active", P["accent_hi"])],
           bordercolor=[("active", P["accent_hi"])], foreground=[("active", P["on_accent"])])
    st.configure("Tool.TButton", background=P["header"], foreground=P["text"], bordercolor=P["header"],
                 lightcolor=P["header"], darkcolor=P["header"], padding=(9, 5), focuscolor=P["header"], width=0)
    st.map("Tool.TButton", background=[("pressed", P["border"]), ("active", P["surface2"])],
           bordercolor=[("active", P["border"])])
    st.configure("ToolAccent.TButton", background=P["accent"], foreground=P["on_accent"], bordercolor=P["accent"],
                 lightcolor=P["accent"], darkcolor=P["accent"], padding=(12, 5), font=F["bold"], focuscolor=P["accent"], width=0)
    st.map("ToolAccent.TButton", background=[("active", P["accent_hi"])], foreground=[("active", P["on_accent"])])
    st.configure("Small.TButton", padding=(7, 3), font=F["small"], width=0)
    # 묶음 툴바 (프로젝트 · 엑셀 · 도구 · 출력 · 편집): 작은 제목 + 테두리 안의 버튼 줄
    st.configure("GroupCap.TLabel", background=P["header"], foreground=P["muted"], font=(fam, size - 2))
    st.configure("Seg.TFrame", background=P["input"], bordercolor=P["border"], relief="solid", borderwidth=1)
    st.configure("Seg.TButton", background=P["input"], foreground=P["text"], bordercolor=P["input"],
                 lightcolor=P["input"], darkcolor=P["input"], padding=(8, 3), focuscolor=P["input"], width=0)
    st.map("Seg.TButton", background=[("pressed", P["border"]), ("active", P["surface2"])],
           bordercolor=[("active", P["surface2"])])
    # 입력 칸
    for w in ("TEntry", "TCombobox", "TSpinbox"):
        st.configure(w, fieldbackground=P["input"], foreground=P["text"], bordercolor=P["border"],
                     lightcolor=P["input"], darkcolor=P["input"], padding=(6, 4), insertcolor=P["text"],
                     arrowcolor=P["muted"], background=P["surface2"])
        st.map(w, bordercolor=[("focus", P["accent"])], lightcolor=[("focus", P["accent"])],
               fieldbackground=[("readonly", P["input"]), ("disabled", P["surface2"])],
               foreground=[("readonly", P["text"]), ("disabled", P["muted"])],
               selectbackground=[("readonly", P["input"])], selectforeground=[("readonly", P["text"])])
    for w in ("TCheckbutton", "TRadiobutton"):
        st.configure(w, background=P["surface"], foreground=P["text"], indicatorbackground=P["input"],
                     indicatorforeground=P["accent"], bordercolor=P["border"], padding=(2, 2), focuscolor=P["surface"])
        st.map(w, background=[("active", P["surface"])], indicatorbackground=[("selected", P["accent"])],
               indicatorforeground=[("selected", P["on_accent"])])
    # 탭
    st.configure("TNotebook", background=P["surface"], bordercolor=P["border"], tabmargins=(0, 0, 0, 0),
                 lightcolor=P["surface"], darkcolor=P["surface"])
    st.configure("TNotebook.Tab", background=P["surface"], foreground=P["muted"], padding=(16, 7),
                 bordercolor=P["surface"], lightcolor=P["surface"], darkcolor=P["surface"], font=F["bold"])
    st.map("TNotebook.Tab", background=[("selected", P["surface2"]), ("active", P["surface2"])],
           foreground=[("selected", P["accent"]), ("active", P["text"])],
           lightcolor=[("selected", P["accent"])], bordercolor=[("selected", P["surface2"])])
    # 표·스크롤바·구분선
    st.configure("Treeview", background=P["surface"], fieldbackground=P["surface"], foreground=P["text"],
                 bordercolor=P["border"], rowheight=26)
    st.map("Treeview", background=[("selected", P["select"])], foreground=[("selected", P["text"])])
    st.configure("Treeview.Heading", background=P["surface2"], foreground=P["muted"], bordercolor=P["border"],
                 relief="flat", font=F["bold"])
    st.configure("TScrollbar", background=P["surface2"], troughcolor=P["surface"], bordercolor=P["surface"],
                 lightcolor=P["surface2"], darkcolor=P["surface2"], arrowcolor=P["muted"], gripcount=0)
    st.map("TScrollbar", background=[("active", P["border"])])
    st.configure("TSeparator", background=P["border"])
    st.configure("TPanedwindow", background=P["bg"])
    st.configure("Sash", sashthickness=10, gripcount=0, background=P["bg"], lightcolor=P["bg"], bordercolor=P["bg"])
    root.configure(background=P["bg"])
    root.option_add("*TCombobox*Listbox.background", P["surface"])
    root.option_add("*TCombobox*Listbox.foreground", P["text"])
    root.option_add("*TCombobox*Listbox.selectBackground", P["accent"])
    root.option_add("*TCombobox*Listbox.selectForeground", P["on_accent"])
    root.option_add("*TCombobox*Listbox.font", F["base"])
    return P, F


def recolor_tk(widget, P, F, keep=()):
    """ttk가 아닌 tk 위젯(목록·글상자·메뉴·창)을 테마 색으로. keep = 그대로 둘 위젯(미리보기 캔버스 등)"""
    import tkinter as tk
    stack = [widget]
    while stack:
        w = stack.pop()
        stack.extend(w.winfo_children())
        if w in keep:
            continue
        try:
            if isinstance(w, (tk.Tk, tk.Toplevel)):
                w.configure(background=P["bg"])
            elif isinstance(w, tk.Listbox):
                w.configure(background=P["surface"], foreground=P["text"], selectbackground=P["select"],
                            selectforeground=P["text"], highlightthickness=0, borderwidth=0,
                            activestyle="none", font=F["list"])
            elif isinstance(w, tk.Text):
                w.configure(background=P["input"], foreground=P["text"], insertbackground=P["text"], font=F["base"],
                            selectbackground=P["select"], highlightthickness=1, highlightbackground=P["border"],
                            highlightcolor=P["accent"], borderwidth=0, relief="flat", padx=6, pady=4)
            elif isinstance(w, tk.Canvas):
                w.configure(background=P["surface"], highlightthickness=0)
            elif isinstance(w, tk.Menu):
                if not IS_MAC:
                    w.configure(background=P["surface"], foreground=P["text"], activebackground=P["accent"],
                                activeforeground=P["on_accent"], borderwidth=0)
        except Exception:
            pass
        if w.winfo_class() == "TCombobox":          # 이미 만들어진 펼침 목록도 색 바꿈
            try:
                pd = w.tk.call("ttk::combobox::PopdownWindow", w)
                w.tk.call(f"{pd}.f.l", "configure", "-background", P["surface"], "-foreground", P["text"],
                          "-selectbackground", P["accent"], "-selectforeground", P["on_accent"])
            except Exception:
                pass


def flow_layout(frame, gap_y=4):
    """frame 안에 pack 해 둔 위젯들을 창 폭에 맞춰 여러 줄로 감싸 배치 (좁은 창에서 버튼이 잘리지 않게)"""
    import tkinter as tk
    items = []
    for w in frame.pack_slaves():
        info = w.pack_info()
        px = info.get("padx", 0)
        px = px if isinstance(px, (tuple, list)) else (px, px)
        items.append((w, int(px[0]), int(px[1]), w.winfo_class() == "TSeparator"))
        w.pack_forget()
    state = {"w": 0}

    def layout(_=None):
        W = frame.winfo_width()
        if W <= 1 or W == state["w"]:
            return
        state["w"] = W
        rowh = max([w.winfo_reqheight() for w, *_ in items if not _[2]] or [24])
        x = y = 0
        for w, l, r, sep in items:
            ww = 1 if sep else w.winfo_reqwidth()
            if x > 0 and x + l + ww + r > W:
                x, y = 0, y + rowh + gap_y
                if sep:                     # 줄 맨 앞의 구분선은 숨김
                    w.place_forget()
                    continue
            if sep:
                w.place(x=x + l, y=y + 4, width=1, height=rowh - 8)
            else:
                w.place(x=x + l, y=y + (rowh - w.winfo_reqheight()) // 2)
            x += l + ww + r
        frame.configure(height=y + rowh)
    frame.bind("<Configure>", layout)
    frame.configure(height=max([w.winfo_reqheight() for w, *_ in items] or [24]))
    return layout


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
    root.geometry("1440x880")
    root.minsize(1100, 680)

    S = {  # 화면 상태
        "routes": [], "sel": None, "photo": None, "loading": False, "project": None, "dirty": False,
        "undo": [], "redo": [], "snap": None, "commit_job": None, "page": 0, "pages": [],
        "cmp_img": None, "cmp_name": "", "last_files": [], "paint_val": None,
    }
    status = tk.StringVar(value="준비")

    # ── 테마 ──
    theme_var = tk.StringVar(value=settings.get("theme") if settings.get("theme") in THEMES else "다크")
    PAL, FON = setup_theme(root, theme_var.get(), ttk)

    # ── 레이아웃: 위 = 도구 막대, 왼쪽 = 노선 목록, 가운데 = 미리보기, 오른쪽 = 편집 패널 ──
    header = ttk.Frame(root, style="Header.TFrame", padding=(14, 8))
    header.pack(fill="x")
    brand = ttk.Frame(header, style="Header.TFrame")
    brand.pack(side="left", padx=(0, 18))
    ttk.Label(brand, text="노선이미지 생성기", style="Brand.TLabel").pack(anchor="w")
    ttk.Label(brand, text=f"BusSign Maker · v{APP_VERSION.split()[0]}", style="BrandSub.TLabel").pack(anchor="w")
    acts = ttk.Frame(header, style="Header.TFrame")       # 오른쪽 끝: 현재 저장 · 전체 생성 (항상 보이게 먼저 배치)
    acts.pack(side="right", padx=(10, 0))
    gen_text = tk.StringVar(value="전체 생성")
    bar = ttk.Frame(header, style="Header.TFrame")
    bar.pack(side="left", fill="both", expand=True)
    ttk.Separator(root, orient="horizontal").pack(fill="x")
    sbar = ttk.Frame(root, style="Status.TFrame")
    sbar.pack(fill="x", side="bottom")
    ttk.Label(sbar, textvariable=status, anchor="w", style="Status.TLabel").pack(side="left", fill="x", expand=True)
    count_var = tk.StringVar(value="")
    ttk.Label(sbar, textvariable=count_var, style="Status.TLabel").pack(side="right")
    ttk.Separator(root, orient="horizontal").pack(fill="x", side="bottom")
    body = ttk.Panedwindow(root, orient="horizontal")
    body.pack(fill="both", expand=True, padx=12, pady=12)

    # 노선 목록 (카드)
    lf = ttk.Frame(body, style="Card.TFrame", padding=12)
    body.add(lf, weight=0)
    lhead = ttk.Frame(lf)
    lhead.pack(fill="x", pady=(0, 8))
    ttk.Label(lhead, text="노선", style="Section.TLabel").pack(side="left")
    lcount = ttk.Label(lhead, text="0", style="Badge.TLabel")
    lcount.pack(side="left", padx=6)
    findf = ttk.Frame(lf)
    findf.pack(fill="x", pady=(0, 8))
    ttk.Label(findf, text="찾기", style="Muted.TLabel").pack(side="left", padx=(0, 6))
    find_var = tk.StringVar()
    find_e = ttk.Entry(findf, textvariable=find_var, width=10)
    find_e.pack(side="left", fill="x", expand=True)
    lbf = ttk.Frame(lf)
    lbf.pack(fill="both", expand=True)
    lb = tk.Listbox(lbf, width=21, exportselection=False)
    lb.pack(side="left", fill="both", expand=True)
    sb = ttk.Scrollbar(lbf, command=lb.yview)
    sb.pack(side="left", fill="y")
    lb.config(yscrollcommand=sb.set)
    lbtn = ttk.Frame(lf)
    lbtn.pack(fill="x", pady=(8, 0))

    # 가운데 = 미리보기 카드 (아래 '미리보기'에서 채움)
    pv = ttk.Frame(body, style="Card.TFrame", padding=12)
    body.add(pv, weight=1)

    # 오른쪽 = 편집 패널 (탭마다 세로 스크롤)
    side = ttk.Frame(body, style="Card.TFrame", padding=(10, 10, 4, 10))
    body.add(side, weight=0)
    mid = ttk.Notebook(side, width=400)
    mid.pack(fill="both", expand=True)
    scrollers = []

    def make_tab(title):
        outer = ttk.Frame(mid)
        cv = tk.Canvas(outer, highlightthickness=0, borderwidth=0)
        vs = ttk.Scrollbar(outer, orient="vertical", command=cv.yview)
        inner = ttk.Frame(cv, padding=(12, 12, 14, 12))
        win = cv.create_window(0, 0, anchor="nw", window=inner)
        cv.configure(yscrollcommand=vs.set)
        vs.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)

        def sync(_=None):
            cv.configure(scrollregion=(0, 0, inner.winfo_reqwidth(), inner.winfo_reqheight()))
            cv.itemconfigure(win, width=cv.winfo_width())
        inner.bind("<Configure>", sync)
        cv.bind("<Configure>", sync)

        def wheel(ev):
            if inner.winfo_reqheight() <= cv.winfo_height():
                return
            d = ev.delta if IS_MAC else ev.delta // 120
            cv.yview_scroll(-d if d else (-1 if getattr(ev, "num", 0) == 4 else 1), "units")
        for w_ in (cv, inner):
            w_.bind("<Enter>", lambda e: (root.bind_all("<MouseWheel>", wheel), root.bind_all("<Button-4>", wheel),
                                          root.bind_all("<Button-5>", wheel)))
            w_.bind("<Leave>", lambda e: (root.unbind_all("<MouseWheel>"), root.unbind_all("<Button-4>"),
                                          root.unbind_all("<Button-5>")))
        mid.add(outer, text=title)
        scrollers.append(cv)
        return inner

    tab_edit, tab_style, tab_font, tab_out = (make_tab(t) for t in ("편집", "스타일", "폰트", "출력"))
    for t_ in (tab_edit, tab_style, tab_font, tab_out):
        t_.columnconfigure(1, weight=1)

    # ── 편집 탭 (묶음별 카드) ──
    vars_ = {}

    def edit_box(title, row):
        f = ttk.LabelFrame(tab_edit, text=title, padding=8)
        f.grid(row=row, column=0, columnspan=2, sticky="we", pady=(0, 8))
        return f

    def ent(parent, key, width=12):
        v = vars_.setdefault(key, tk.StringVar())
        return ttk.Entry(parent, textvariable=v, width=width)

    bf_ = edit_box("기본 정보", 0)
    bf_.columnconfigure(1, weight=1)
    bf_.columnconfigure(3, weight=1)
    ttk.Label(bf_, text="번호").grid(row=0, column=0, sticky="e", padx=(0, 4), pady=2)
    ent(bf_, "no", 10).grid(row=0, column=1, sticky="we", pady=2)
    ttk.Label(bf_, text="구분").grid(row=0, column=2, sticky="e", padx=(10, 4))
    kind_var = tk.StringVar(value="일반")
    ttk.Combobox(bf_, textvariable=kind_var, values=KINDS, state="readonly", width=7).grid(
        row=0, column=3, sticky="we", pady=2)
    ttk.Label(bf_, text="템플릿").grid(row=1, column=0, sticky="e", padx=(0, 4))
    tpl_var = tk.StringVar(value="jinmyeong_front_192x160")
    tpl_cb = ttk.Combobox(bf_, textvariable=tpl_var, values=all_template_names(), state="readonly", width=28)
    tpl_cb.grid(row=1, column=1, columnspan=3, sticky="we", pady=2)

    sef = edit_box("기점 · 종점", 1)
    sef.columnconfigure(1, weight=1)
    sef.columnconfigure(2, weight=1)
    ttk.Label(sef, text="기점", style="Muted.TLabel").grid(row=0, column=1, sticky="w")
    ttk.Label(sef, text="종점", style="Muted.TLabel").grid(row=0, column=2, sticky="w", padx=(6, 0))
    for i, (lang, a_, b_) in enumerate([("한글", "start_ko", "end_ko"), ("영어", "start_en", "end_en"),
                                        ("日本", "start_ja", "end_ja"), ("中文", "start_zh", "end_zh")]):
        ttk.Label(sef, text=lang).grid(row=i + 1, column=0, sticky="e", padx=(0, 6), pady=2)
        ent(sef, a_).grid(row=i + 1, column=1, sticky="we", pady=2)
        ent(sef, b_).grid(row=i + 1, column=2, sticky="we", pady=2, padx=(6, 0))
    ttk.Label(sef, text="'김해|구산동'처럼 | 를 넣으면 두 줄 화면에서 그 자리에서 줄을 바꿉니다. "
                        "日·中을 비워 두면 사전에서 찾아 넣습니다.",
              style="Muted.TLabel", wraplength=320, justify="left").grid(row=5, column=0, columnspan=3, sticky="w",
                                                                        pady=(4, 0))

    vf_ = edit_box("경유지", 2)
    vf_.columnconfigure(0, weight=1)
    via_txt = tk.Text(vf_, width=30, height=4, wrap="word", undo=False)
    via_txt.grid(row=0, column=0, sticky="we")
    ttk.Label(vf_, text="쉼표 = 정류장, 줄바꿈 = 줄 (한글 줄 → 영문 줄)", style="Muted.TLabel").grid(
        row=1, column=0, sticky="w", pady=(4, 0))

    xf = edit_box("추가 칸 — 부산 구형 · 신형 · 풀컬러", 3)
    xf.columnconfigure(1, weight=1)
    xf.columnconfigure(3, weight=1)
    ttk.Label(xf, text="문구").grid(row=0, column=0, sticky="e", padx=(0, 4), pady=2)
    ent(xf, "slogan").grid(row=0, column=1, sticky="we", pady=2)
    ttk.Label(xf, text="잔여좌석").grid(row=0, column=2, sticky="e", padx=(10, 4))
    ent(xf, "seats", 5).grid(row=0, column=3, sticky="we", pady=2)
    ttk.Label(xf, text="로고").grid(row=1, column=0, sticky="e", padx=(0, 4))
    ent(xf, "logo").grid(row=1, column=1, columnspan=2, sticky="we", pady=2)

    def pick_logo():
        p = filedialog.askopenfilename(title="로고 이미지 (PNG·BMP·JPG)",
                                       filetypes=[("이미지", "*.png *.bmp *.jpg *.jpeg *.gif"), ("모든 파일", "*.*")])
        if p:
            vars_["logo"].set(p)

    ttk.Button(xf, text="찾기…", style="Small.TButton", command=pick_logo).grid(row=1, column=3, sticky="w", padx=(6, 0))
    ttk.Label(xf, text="로고는 판 크기에 맞춰 줄여 한 장면으로 넣습니다(구형은 3색으로 바꿈). "
                       "잔여좌석에 숫자를 넣으면 '잔여좌석 24' 화면이 추가됩니다.",
              style="Muted.TLabel", wraplength=320, justify="left").grid(row=2, column=0, columnspan=4, sticky="w",
                                                                        pady=(4, 0))
    rom_frame = edit_box("영문 자동 변환 (국어의 로마자 표기법 + 사전)", 4)
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
                  else list(WEIGHTS) if k == "weight" else list(NUM_WEIGHTS) if k == "num_weight"
                  else OLD_BOX_STYLES if k == "old_box" else OLD_GRAD_STYLES if k == "old_grad"
                  else OLD_ARROW_STYLES if k == "old_arrow" else NEW_VIA_STYLES if k == "new_via"
                  else NEW_EXPRESS_STYLES if k == "new_express" else FC_NUM_STYLES if k == "fc_num"
                  else FC_DIR_STYLES if k == "fc_dir" else FC_SEP_STYLES if k == "fc_sep" else list(COLORS))
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
              style="Muted.TLabel", justify="left").grid(row=3, column=0, columnspan=3, sticky="w", pady=(8, 0))

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
              style="Muted.TLabel", justify="left", wraplength=300).grid(row=2, column=1, columnspan=2, sticky="w")
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
    phead = ttk.Frame(pv)
    phead.pack(fill="x", pady=(0, 8))
    ttk.Label(phead, text="미리보기", style="Section.TLabel").pack(side="left")
    tpl_badge = ttk.Label(phead, text="", style="Badge.TLabel")
    tpl_badge.pack(side="left", padx=8)
    pctl = ttk.Frame(pv)
    pctl.pack(fill="x")
    zoom_var = tk.StringVar(value=str(settings.get("zoom", "맞춤")))
    ttk.Label(pctl, text="배율").pack(side="left")
    ttk.Combobox(pctl, textvariable=zoom_var, values=["맞춤", 2, 3, 4, 5, 6, 8], width=4, state="readonly").pack(side="left", padx=(2, 8))
    page_var = tk.StringVar(value="1")
    ttk.Label(pctl, text="페이지").pack(side="left")
    page_cb = ttk.Combobox(pctl, textvariable=page_var, values=["1"], width=3, state="readonly")
    page_cb.pack(side="left", padx=(2, 8))
    ttk.Separator(pctl, orient="vertical").pack(side="left", fill="y", padx=6)
    led_var = tk.BooleanVar(value=bool(settings.get("led_look", True)))
    ttk.Checkbutton(pctl, text="LED 모양", variable=led_var).pack(side="left", padx=(2, 0))
    play_btn = ttk.Button(pctl, text="▶ 재생", width=7, style="Accent.TButton")
    play_btn.pack(side="left", padx=(10, 0))
    pctl2 = ttk.Frame(pv)                      # 둘째 줄: 점 편집 · 원본 비교
    pctl2.pack(fill="x", pady=(6, 0))
    mode_var = tk.StringVar(value="보기")
    ttk.Radiobutton(pctl2, text="보기", variable=mode_var, value="보기").pack(side="left")
    ttk.Radiobutton(pctl2, text="점 편집", variable=mode_var, value="편집").pack(side="left", padx=(6, 0))
    ttk.Label(pctl2, text="펜", style="Muted.TLabel").pack(side="left", padx=(10, 2))
    pen_var = tk.StringVar(value="흰색")
    ttk.Combobox(pctl2, textvariable=pen_var, values=PEN_COLORS, width=9, state="readonly").pack(side="left")
    ttk.Separator(pctl2, orient="vertical").pack(side="left", fill="y", padx=10)
    cmp_var = tk.BooleanVar(value=False)
    ttk.Checkbutton(pctl2, text="원본 겹쳐보기", variable=cmp_var).pack(side="left")
    flow_layout(pctl)
    flow_layout(pctl2)
    cwrap = ttk.Frame(pv)
    cwrap.pack(fill="both", expand=True, pady=(10, 0))
    canvas = tk.Canvas(cwrap, background=PAL["canvas"], highlightthickness=0)
    hsb = ttk.Scrollbar(cwrap, orient="horizontal", command=canvas.xview)
    vsb = ttk.Scrollbar(cwrap, orient="vertical", command=canvas.yview)
    canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
    vsb.pack(side="right", fill="y")
    hsb.pack(side="bottom", fill="x")
    canvas.pack(side="left", fill="both", expand=True)
    img_item = canvas.create_image(0, 0, anchor="nw")
    # 쪽 썸네일 줄: 모든 쪽을 작게 늘어놓고, 누르면 그 쪽으로 이동 (쪽 종류 · 표시 시간 표시)
    STRIP_H, THUMB_H, THUMB_W = 70, 36, 150
    strip_wrap = ttk.Frame(pv)
    strip_wrap.pack(fill="x", pady=(8, 0))
    strip = tk.Canvas(strip_wrap, height=STRIP_H, highlightthickness=0, background=PAL["canvas"])
    strip_sb = ttk.Scrollbar(strip_wrap, orient="horizontal", command=strip.xview)
    strip.configure(xscrollcommand=strip_sb.set)
    strip.pack(fill="x")
    strip_sb.pack(fill="x")
    PAGE_KIND_LABEL = {"big": "큰 번호", "route": "행선", "via": "경유지", "seats": "잔여좌석", "slogan": "문구",
                       "logo": "로고"}
    S["strip_key"], S["thumbs"], S["thumb_boxes"] = None, [], []

    def draw_strip():
        import hashlib
        pages = S.get("pages") or []
        key = (tuple(hashlib.md5(p.tobytes()).hexdigest() for p in pages), bool(led_var.get()), PAL["canvas"])
        if S["strip_key"] != key:
            S["strip_key"] = key
            strip.delete("all")
            S["thumbs"], S["thumb_boxes"] = [], []
            x = 10
            for i, p in enumerate(pages):
                pr = p.convert("RGB")
                sc = min(THUMB_H / pr.height, THUMB_W / pr.width)
                tw, th = max(1, round(pr.width * sc)), max(1, round(pr.height * sc))
                if led_var.get() and sc >= 0.9:
                    im = led_look(pr, max(2, int(sc) + 1)).resize((tw, th), Image.LANCZOS)
                else:
                    im = pr.resize((tw, th), Image.NEAREST if sc >= 1 else Image.BOX)
                ph = ImageTk.PhotoImage(im)
                S["thumbs"].append(ph)
                tag = f"pg{i}"
                y0 = 8 + (THUMB_H - th) // 2
                strip.create_image(x, y0, anchor="nw", image=ph, tags=(tag,))
                kind = PAGE_KIND_LABEL.get(p.info.get("page_kind"), "")
                dur = p.info.get("duration")
                cap = f"{i + 1} {kind}".strip() + (f" · {dur / 1000:.1f}초" if dur and len(pages) > 1 else "")
                strip.create_text(x, 8 + THUMB_H + 6, anchor="nw", text=cap, fill="#b4bac6", font=FON["small"],
                                  tags=(tag,))
                S["thumb_boxes"].append((x - 4, 4, x + tw + 4, 8 + THUMB_H + 4))
                strip.tag_bind(tag, "<Button-1>", lambda e, i=i: page_var.set(str(i + 1)))
                x += max(tw, 88) + 12
            strip.configure(scrollregion=(0, 0, x, STRIP_H))
        strip.delete("sel")
        boxes = S["thumb_boxes"]
        if boxes and 0 <= S["page"] < len(boxes):
            bx = boxes[S["page"]]
            strip.create_rectangle(*bx, outline=PAL["accent"], width=2, tags=("sel",))
            total = max(1, boxes[-1][2] + 16)
            vw = max(1, strip.winfo_width())
            left, right = strip.canvasx(0), strip.canvasx(vw)
            if bx[0] < left or bx[2] > right:              # 고른 쪽이 보이도록 줄을 넘김
                strip.xview_moveto(max(0.0, (bx[0] - 20) / total))

    info = tk.StringVar()
    warn_var = tk.StringVar()
    info_lbl = ttk.Label(pv, textvariable=info, style="Muted.TLabel")
    info_lbl.pack(anchor="w", pady=(8, 0))
    warn_lbl = ttk.Label(pv, textvariable=warn_var, style="Warn.TLabel", wraplength=760, justify="left")
    warn_lbl.pack(anchor="w")
    pv.bind("<Configure>", lambda e: warn_lbl.configure(wraplength=max(200, e.width - 30)), add="+")
    legend = ttk.Label(pv, text="", style="Muted.TLabel")
    legend.pack(anchor="w")
    # 아래쪽 줄(썸네일·정보·경고)을 먼저 아래에서부터 배치 → 작은 창에서도 잘리지 않고, 남는 높이는 미리보기 칸이 씀
    for w_ in (cwrap, strip_wrap, info_lbl, warn_lbl, legend):
        w_.pack_forget()
    legend.pack(side="bottom", anchor="w")
    warn_lbl.pack(side="bottom", anchor="w")
    info_lbl.pack(side="bottom", anchor="w", pady=(6, 0))
    strip_wrap.pack(side="bottom", fill="x", pady=(8, 0))
    cwrap.pack(side="top", fill="both", expand=True, pady=(10, 0))

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
        s_ = (r.get("start_ko") or "").replace("|", "").strip()
        e_ = (r.get("end_ko") or "").replace("|", "").strip()
        path = f"{s_}→{e_}" if s_ and e_ else (e_ or s_)
        return (r.get("no") or "(새 노선)") + (f"  {path}" if path else "") + (f"  [{k}]" if k and k != "일반" else "")

    def refill_list(keep=None):
        lcount.config(text=str(len(S["routes"])))
        count_var.set(f"노선 {len(S['routes'])}개")
        gen_text.set(f"전체 생성 · {len(S['routes'])}개" if S["routes"] else "전체 생성")
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
        if S.get("play_job"):
            stop_play()
        r = cur()
        if r is None:
            canvas.itemconfig(img_item, image="")
            info.set("노선이 없습니다. [새 노선] 또는 엑셀/프로젝트를 불러오세요.")
            S["pages"] = []
            draw_strip()
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
        if led_var.get() and mode_var.get() != "편집" and not (cmp_var.get() and S["cmp_img"] is not None) and z >= 2:
            big = led_look(img, z)                    # 사진처럼: 둥근 LED 점 + 빛 번짐
        else:
            big = img.resize((img.width * z, img.height * z), Image.NEAREST)
        if mode_var.get() == "편집" and z >= 3:     # 편집할 때는 격자를 얇게 표시
            d = ImageDraw.Draw(big)
            for x in range(0, big.width, z):
                d.line([(x, 0), (x, big.height)], fill=(40, 40, 40))
            for y in range(0, big.height, z):
                d.line([(0, y), (big.width, y)], fill=(40, 40, 40))
        S["photo"] = ImageTk.PhotoImage(big)
        canvas.itemconfig(img_item, image=S["photo"])
        ox = max(0, (canvas.winfo_width() - big.width) // 2)          # 가운데에 놓기
        oy = max(0, (canvas.winfo_height() - big.height) // 2)
        S["off"] = (ox, oy)
        canvas.coords(img_item, ox, oy)
        canvas.configure(scrollregion=(0, 0, max(canvas.winfo_width(), big.width + ox), max(canvas.winfo_height(), big.height + oy)))
        tpl_badge.config(text=f"{r.get('template') or ''}  ·  {S['page'] + 1}/{n}쪽")
        names = file_names(r, n, settings.get("name_pattern"))
        info.set(f"{img.width}x{img.height} px  |  {r.get('template')}  |  {names[S['page']]}"
                 f"{'  |  편집한 점 ' + str(len((r.get('edits') or {}).get(str(S['page']), {}))) + '개' if (r.get('edits') or {}).get(str(S['page'])) else ''}{extra}")
        draw_strip()
        ws = [w for w in S.get("warns", []) if w[1] > 0]
        warn_var.set("⚠ 압축됨: " + " / ".join(f"{t} ({ratio * 100:.0f}%)" for t, ratio in ws) if ws else "")

    S["play_job"] = None

    def stop_play():
        if S.get("play_job"):
            root.after_cancel(S["play_job"])
        S["play_job"] = None
        play_btn.config(text="▶ 재생")

    def toggle_play():
        if S.get("play_job"):
            stop_play(); refresh_preview(); return
        pages = S.get("pages") or []
        if not pages:
            return
        frames = anim_frames(pages)
        z = zoom_now(pages[0].width, pages[0].height)
        led = led_var.get()
        cache = {}
        play_btn.config(text="■ 정지")

        def step(i):
            f, ms = frames[i % len(frames)]
            key = i % len(frames)
            if key not in cache:
                cache[key] = ImageTk.PhotoImage(led_look(f, z) if led else f.resize((f.width * z, f.height * z), Image.NEAREST))
            canvas.itemconfig(img_item, image=cache[key])
            S["play_job"] = root.after(ms, lambda: step(i + 1))

        step(0)

    play_btn.config(command=toggle_play)

    def save_gif_cmd():
        r = cur()
        if r is None:
            return
        p = filedialog.asksaveasfilename(defaultextension=".gif", initialfile=f"{safe_name(r.get('no') or 'noname')}.gif",
                                         filetypes=[("GIF", "*.gif")])
        if not p:
            return
        pages = render_pages(r, current_style())
        z = 4 if pages[0].width <= 192 else 3
        n = save_gif(pages, p, z=z, led=led_var.get())
        status.set(f"GIF 저장: {p} ({n}프레임, LED 모양 {'켬' if led_var.get() else '끔'})")

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
        ox, oy = S.get("off", (0, 0))
        x = int((canvas.canvasx(ev.x) - ox) // z)
        y = int((canvas.canvasy(ev.y) - oy) // z)
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

    def find_route(step=0):
        """번호·지명(한글/영문)으로 노선 찾기. 엔터 = 다음 것"""
        q = find_var.get().strip().lower()
        if not q or not S["routes"]:
            return
        n = len(S["routes"])
        start = ((S["sel"] or 0) + step) % n
        for k in range(n):
            i = (start + k) % n
            r = S["routes"][i]
            hay = " ".join([label_of(r), r.get("start_en", ""), r.get("end_en", ""), r.get("via", "")]).lower()
            if q in hay:
                if i != S["sel"]:
                    if S["commit_job"]:
                        root.after_cancel(S["commit_job"]); commit()
                    refill_list(i)
                    load_form(i)
                status.set(f"찾음: {label_of(r)}  (엔터 = 다음)")
                return
        status.set(f"'{find_var.get().strip()}' 노선을 찾지 못했습니다")

    find_var.trace_add("write", lambda *_: find_route(0))
    find_e.bind("<Return>", lambda e: find_route(1))
    ttk.Button(lbtn, text="＋ 새 노선", style="Accent.TButton", command=new_route).pack(fill="x", pady=(0, 6))
    lrow = ttk.Frame(lbtn)
    lrow.pack(fill="x")
    for t, c in [("복제", dup_route), ("삭제", del_route), ("▲", lambda: move_route(-1)), ("▼", lambda: move_route(1))]:
        ttk.Button(lrow, text=t, style="Small.TButton", width=4 if len(t) > 1 else 2,
                   command=c).pack(side="left", padx=(0, 4), expand=len(t) > 1, fill="x")

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

    ttk.Button(rf2, text="이 노선 채우기", style="Accent.TButton", command=romanize_current).pack(side="left")
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
    ttk.Button(pf_, text="프리셋 저장", style="Accent.TButton", command=save_preset).pack(side="left", padx=6)

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
        w.geometry("680x420")
        w.configure(background=PAL["surface"])
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
        ttk.Button(bf, text="닫기", style="Accent.TButton", command=w.destroy).pack(side="right")

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
        w.geometry("940x660")
        w.configure(background=PAL["surface"])
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
        ttk.Label(right, textvariable=msg, style="Muted.TLabel").pack(anchor="w")
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

        ttk.Button(bot, text="저장", style="Accent.TButton", command=save_glyph).pack(side="left")
        ttk.Button(bot, text="기본값으로 되돌리기", command=reset_glyph).pack(side="left", padx=6)
        ttk.Button(bot, text="모두 지우기", command=lambda: (G["m"].fill(False), draw()) if G["m"] is not None else None).pack(side="left")
        ttk.Button(bot, text="닫기", command=w.destroy).pack(side="right")
        refresh_list()
        load_glyph()
        recolor_tk(w, PAL, FON, keep=(gc,))

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
    m_out.add_separator()
    m_out.add_command(label="움직이는 GIF로 저장 (LED 모양)…", command=save_gif_cmd)
    menubar.add_cascade(label="출력", menu=m_out)
    root.config(menu=menubar)

    def tool_group(title, items):
        """작은 제목 + 테두리 안 버튼 줄. 좁은 창에서는 flow_layout 이 묶음째 다음 줄로 넘김"""
        g_ = ttk.Frame(bar, style="Header.TFrame")
        ttk.Label(g_, text=title, style="GroupCap.TLabel").pack(anchor="w", padx=(3, 0))
        row_ = ttk.Frame(g_, style="Seg.TFrame", padding=2)
        row_.pack(anchor="w")
        for text, cmd in items:
            ttk.Button(row_, text=text, command=cmd, style="Seg.TButton",
                       width=3 if len(text) == 1 else None).pack(side="left", padx=1)
        g_.pack(side="left", padx=(0, 10))
        return row_

    tool_group("프로젝트", [("열기", open_project), ("저장", save_project_cmd)])
    tool_group("엑셀", [("불러오기", load_xl), ("내보내기", export_xl)])
    tool_group("도구", [("글자 보정", glyph_editor), ("원본 비교", load_compare), ("영문 채우기", romanize_current)])
    tool_group("출력", [("출력 폴더", lambda: open_folder(out_dir())), ("USB 복사", copy_usb), ("GIF", save_gif_cmd)])
    tool_group("편집", [("↶", undo), ("↷", redo)])
    view_row = tool_group("보기", [])
    ttk.Button(acts, text="현재 저장 (BMP)", command=save_current, style="Tool.TButton").pack(side="left", padx=(0, 6))
    gen_btn = ttk.Button(acts, textvariable=gen_text, command=save_all, style="ToolAccent.TButton")
    gen_btn.pack(side="left")
    gen_text.set(f"전체 생성 · {len(S['routes'])}개" if S["routes"] else "전체 생성")

    def apply_theme(*_):
        nonlocal PAL, FON
        PAL, FON = setup_theme(root, theme_var.get(), ttk)
        recolor_tk(root, PAL, FON, keep=(canvas, strip))
        for c_ in scrollers:
            c_.configure(background=PAL["surface"])
        canvas.configure(background=PAL["canvas"])
        strip.configure(background=PAL["canvas"])
        draw_strip()
        theme_btn.config(text="☀ 라이트" if theme_var.get() == "다크" else "☾ 다크")
        settings["theme"] = theme_var.get()
        try:
            save_settings(settings)
        except Exception:
            pass

    theme_btn = ttk.Button(view_row, text="☀ 라이트", style="Seg.TButton",
                           command=lambda: theme_var.set("라이트" if theme_var.get() == "다크" else "다크"))
    theme_btn.pack(side="left", padx=1)
    flow_layout(bar)
    theme_var.trace_add("write", apply_theme)
    m_view = tk.Menu(menubar, tearoff=0)
    for t_ in THEMES:
        m_view.add_radiobutton(label=f"{t_} 테마", variable=theme_var, value=t_)
    m_view.add_separator()
    m_view.add_checkbutton(label="LED 모양으로 미리보기", variable=led_var)
    menubar.insert_cascade(2, label="보기", menu=m_view)
    apply_theme()

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
    led_var.trace_add("write", lambda *_: (settings.update(led_look=bool(led_var.get())), refresh_preview()))
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
    {"no": "37", "start_ko": "금정공영차고지", "end_ko": "고원아파트", "start_en": "Geumjeong Bus Garage",
     "end_en": "Gowon APT.", "start_ja": "金井コン ヨン・チャゴジ", "end_ja": "Gowonアパート",
     "start_zh": "金井 公營車庫", "end_zh": "Gowon公寓",
     "via": "노포동터미널,월평삼거리\nNopo-dong Bus Term.,Institute of radiological & medical sciences",
     "template": "busan_old_192x32 (부산 구형 3색)"},
    {"no": "99", "start_ko": "남포동", "end_ko": "부산진시장", "start_en": "Nampo-dong", "end_en": "Busanjin Market",
     "via": "충무교차로,시청,서면,범내골\nChungmu Intersection,City Hall,Seomyeon,Beomnaegol",
     "seats": "24", "slogan": "Busan is good", "template": "busan_new_192x32 (부산 신형 풀컬러)"},
    {"no": "2029", "start_ko": "녹산", "end_ko": "해운대", "start_en": "Noksan", "end_en": "Haeundae", "kind": "급행",
     "via": "명지,하단,괴정,서면,센텀시티\nMyeongji,Hadan,Goejeong,Seomyeon,Centum City",
     "slogan": "Busan is good", "template": "busan_new_192x32 (부산 신형 풀컬러)"},
    {"no": "148-1", "start_ko": "학장동", "end_ko": "부산대학교", "seats": "12",
     "template": "fullcolor_dir_192x32 (풀컬러 방향)"},
]
SELFTEST_STYLES = [{}, {"num_weight": "균일 2픽셀", "arrow_shape": "→"}, {"num_weight": "균일 3픽셀"}]
SELFTEST_EXPECTED = {'360#0': '4dd240bfda9c43c2', '1100#0': 'a103fbb8f3c73b85', '58-1#0': '1bcf005ba43f14e2', '7-12#0': '44a7a4a1346c97ff', '37#0': '95a7f5d972e3aead', '99#0': 'ecea271bbd92a6b4', '2029#0': '01386d76aaf6f832', '148-1#0': 'bd0e0213f905332c', '360#1': 'a30ee2e589538682', '1100#1': 'c0e38276c8232e1b', '58-1#1': '87bafe9b72384f83', '7-12#1': '5fe5c322428ba176', '37#1': '95a7f5d972e3aead', '99#1': 'ecea271bbd92a6b4', '2029#1': '01386d76aaf6f832', '148-1#1': 'bd0e0213f905332c', '360#2': '15b0ac3bfbf10f6f', '1100#2': '50735154bb7b0e58', '58-1#2': '79e4b8cf9575fe6a', '7-12#2': '44a7a4a1346c97ff', '37#2': '95a7f5d972e3aead', '99#2': 'ecea271bbd92a6b4', '2029#2': '01386d76aaf6f832', '148-1#2': 'bd0e0213f905332c'}  # 검증 완료된 기준 이미지 해시


def selftest_hashes():
    import hashlib
    out = {}
    saved = dict(GLYPHS)
    GLYPHS.clear()
    for si, st in enumerate(SELFTEST_STYLES):
        for r in SELFTEST_ROUTES:
            img = render_route(r, st)
            data = b"".join(p.tobytes() for p in img) if isinstance(img, list) else img.tobytes()
            out[f"{r['no']}#{si}"] = hashlib.sha256(data).hexdigest()[:16]
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
    # 창 모드로 묶인 윈도우 exe 에는 콘솔이 없어서 print 가 실패함 → 버리는 곳으로 돌림
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")
    try:
        sys.stdout.reconfigure(errors="replace")   # 한글을 못 쓰는 콘솔에서도 멈추지 않게
    except Exception:
        pass
    if len(sys.argv) > 1 and sys.argv[1] == "--selftest":
        sys.exit(run_selftest())
    if len(sys.argv) > 2 and sys.argv[1] == "--render":  # 테스트용: --render out.bmp
        r = {"no": "360", "start_ko": "서울역 환승센터", "end_ko": "강일IC",
             "start_en": "Seoul Station", "end_en": "Gang-il IC",
             "via": "동묘앞,을지로4가,\n강동역,십자성마을\nDongmyo,Euljuro-4ga,\nGangdong Stn,10JASUNG TOWN"}
        render_route(r).save(sys.argv[2], "BMP")
    else:
        run_gui()
