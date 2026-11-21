#!/bin/bash
# ─────────────────────────────────────────────
#  노선이미지 생성기 → BusSignMaker.app / .dmg 빌드
#  더블클릭 또는 터미널에서:  bash build_dmg.command
# ─────────────────────────────────────────────
set -e
cd "$(dirname "$0")"
APP="BusSignMaker"
DISPLAY_NAME="노선이미지 생성기"
VERSION="2.0"

echo "▶ 1/5  tkinter 되는 파이썬 찾는 중…"
PY=""
CANDS=(/Library/Frameworks/Python.framework/Versions/3.*/bin/python3 /opt/homebrew/bin/python3 /usr/local/bin/python3 python3)
for c in "${CANDS[@]}"; do
  if command -v "$c" >/dev/null 2>&1; then
    # Tk 8.6 이상만 사용 (애플 기본 python3의 Tk 8.5는 화면이 깨짐)
    if "$c" -c "import tkinter,sys; v=tkinter.Tcl().eval('info patchlevel'); sys.exit(0 if tuple(map(int,v.split('.')[:2]))>=(8,6) else 1)" 2>/dev/null; then
      PY="$c"; break
    fi
  fi
done
if [ -z "$PY" ]; then
  echo ""
  echo "✖ tkinter(Tk 8.6+)가 들어있는 파이썬이 없습니다. 둘 중 하나를 설치한 뒤 다시 실행하세요:"
  echo "   · python.org 에서 macOS용 Python 3.12 설치 (가장 간단)"
  echo "   · 또는 Homebrew:  brew install python python-tk"
  read -n 1 -s -r -p "아무 키나 누르면 닫힙니다…"; exit 1
fi
echo "   → $PY ($("$PY" --version))"

echo "▶ 2/5  빌드용 가상환경 준비…"
"$PY" -m venv .build_venv
source .build_venv/bin/activate
pip install -q --upgrade pip
# 검증한 것과 똑같은 버전으로 고정 (버전이 다르면 글자 픽셀이 달라질 수 있음)
pip install -q pyinstaller "pillow==12.3.0" "numpy==2.5.3" "openpyxl==3.1.5" || {
  echo "✖ 고정 버전 설치 실패 — 파이썬 버전($("$PY" --version))용 패키지가 없을 수 있습니다. Python 3.12 또는 3.13을 쓰세요."; exit 1; }

echo "   · 소스 상태로 먼저 자체검사…"
python bus_sign_maker.py --selftest || { echo "✖ 자체검사 실패 — 이 결과를 캡처해서 보내주세요."; exit 1; }

echo "▶ 3/5  아이콘 만드는 중…"
rm -rf .icon.iconset && mkdir .icon.iconset
for s in 16 32 128 256 512; do
  sips -z $s $s icon.png --out .icon.iconset/icon_${s}x${s}.png >/dev/null
  d=$((s*2)); sips -z $d $d icon.png --out .icon.iconset/icon_${s}x${s}@2x.png >/dev/null
done
iconutil -c icns .icon.iconset -o icon.icns

echo "▶ 4/5  앱 빌드 (1~2분)…"
rm -rf build dist
pyinstaller --noconfirm --clean --windowed \
  --name "$APP" --icon icon.icns \
  --osx-bundle-identifier "com.guk.bussignmaker" \
  --add-data "fonts:fonts" \
  --hidden-import numpy \
  --hidden-import PIL._tkinter_finder \
  --hidden-import PIL.ImageTk \
  bus_sign_maker.py >/dev/null 2>&1 || { echo "✖ PyInstaller 실패 — 아래 명령으로 로그 확인:"; echo "  source .build_venv/bin/activate && pyinstaller --windowed bus_sign_maker.py"; exit 1; }

PL="dist/$APP.app/Contents/Info.plist"
plutil -replace CFBundleDisplayName -string "$DISPLAY_NAME" "$PL"
plutil -replace CFBundleShortVersionString -string "$VERSION" "$PL"
plutil -replace NSHighResolutionCapable -bool true "$PL"
plutil -replace LSApplicationCategoryType -string "public.app-category.utilities" "$PL"
codesign --force --deep -s - "dist/$APP.app" >/dev/null 2>&1 || true
xattr -cr "dist/$APP.app"

echo "   · 완성된 앱으로 자체검사…"
"dist/$APP.app/Contents/MacOS/$APP" --selftest || { echo "✖ 앱 자체검사 실패 — 이 결과를 캡처해서 보내주세요."; exit 1; }

echo "▶ 5/5  DMG 만드는 중…"
STAGE=".dmg_stage"; rm -rf "$STAGE"; mkdir "$STAGE"
cp -R "dist/$APP.app" "$STAGE/"
ln -s /Applications "$STAGE/Applications"
rm -f "$APP.dmg"
hdiutil create -volname "$DISPLAY_NAME" -srcfolder "$STAGE" -ov -format UDZO "$APP.dmg" >/dev/null
rm -rf "$STAGE" .icon.iconset build "$APP.spec"
deactivate

echo ""
echo "✔ 완료!"
echo "   앱 : $(pwd)/dist/$APP.app"
echo "   DMG: $(pwd)/$APP.dmg"
open -R "$APP.dmg"
