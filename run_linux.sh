#!/bin/bash
# 리눅스에서 소스로 실행: 처음 한 번 가상환경을 만들고, 그다음부터는 바로 실행합니다.
#   bash run_linux.sh            → 프로그램 실행
#   bash run_linux.sh --selftest → 자체검사
set -e
cd "$(dirname "$0")"
if ! python3 -c "import tkinter" 2>/dev/null; then
  echo "✖ tkinter가 없습니다. 먼저 설치하세요:"
  echo "   Ubuntu/Debian : sudo apt install python3-tk python3-venv"
  echo "   Fedora        : sudo dnf install python3-tkinter"
  echo "   Arch          : sudo pacman -S tk"
  exit 1
fi
if [ ! -d .venv ]; then
  echo "▶ 처음 실행: 가상환경 만드는 중…"
  python3 -m venv .venv
  .venv/bin/pip install -q --upgrade pip
  .venv/bin/pip install -q "pillow==12.3.0" "numpy==2.5.3" "openpyxl==3.1.5"
fi
exec .venv/bin/python bus_sign_maker.py "$@"
