## 내려받기

| 운영체제 | 파일 |
|---|---|
| 윈도우 10/11 | `BusSignMaker-Windows.exe` (설치 없이 바로 실행) |
| 맥 (애플 실리콘 M1~M4) | `BusSignMaker-macOS-AppleSilicon.dmg` |
| 리눅스 (x86_64, Ubuntu 22.04 이후 등) | `BusSignMaker-Linux-x86_64.tar.gz` |

세 파일 모두 GitHub 서버에서 자동으로 빌드했고, 빌드할 때 기준 이미지 24세트와 픽셀 단위로 같은지 자체검사를 통과한 것만 올라갑니다.

## 처음 실행할 때

개인이 만든 프로그램이라 인증서 서명이 없어서, 처음 한 번은 경고가 나옵니다.

**윈도우**: "Windows의 PC 보호" 창이 뜨면 **추가 정보 → 실행**을 누르세요.

**맥**: DMG를 열고 앱을 응용 프로그램 폴더로 끌어다 놓은 뒤,
- 앱을 **우클릭 → 열기 → 열기**를 누르세요.
- "손상되었기 때문에 열 수 없습니다"라고 나오면 터미널에서 아래를 한 번 실행하세요.
  ```
  xattr -dr com.apple.quarantine /Applications/BusSignMaker.app
  ```

**리눅스**: 압축을 풀고 실행하세요.
```
tar xzf BusSignMaker-Linux-x86_64.tar.gz
./BusSignMaker-Linux-x86_64/BusSignMaker
```

## 엑셀 샘플

[`samples/노선샘플.xlsx`](https://github.com/M32k/bus-sign-maker/raw/main/samples/%EB%85%B8%EC%84%A0%EC%83%98%ED%94%8C.xlsx)를 받아서 [파일 → 엑셀 불러오기]로 열어 보세요. 부산 구형 3색·신형 풀컬러(잔여좌석 화면 포함), 급행·심야·마을버스, 측면·후면판, 영문 자동 변환, 압축 경고 예시가 들어 있습니다.

## 사용법

[README](https://github.com/M32k/bus-sign-maker#readme)를 참고하세요. 출력 BMP, 프리셋, 글자 보정표 등은 `내 폴더/BusSignMaker` 에 저장됩니다.
