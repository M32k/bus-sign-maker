"""부산 · 서울 실제 노선 샘플 엑셀 만들기 (samples/부산노선_샘플.xlsx, samples/서울노선_샘플.xlsx)

노선 정보(기점·종점·주요 정류장)는 2026-11-26에 부산 BIMS(bus.busan.go.kr)·서울 버스정보시스템(bus.go.kr)의
노선별 정류소 목록과 대조한 내용입니다(서울 N26의 경유 정류장만 보조 출처 기준).
노선은 바뀔 수 있으니 실제로 쓸 때는 운수사·지자체 최신 정보로 확인하세요.
실행:  python3 tools/make_city_samples.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
import bus_sign_maker as b  # noqa: E402

OLD, NEW, FC = "busan_old_192x32 (부산 구형 3색)", "busan_new_192x32 (부산 신형 풀컬러)", "fullcolor_dir_192x32 (풀컬러 방향)"
JF, JU, SF = "jinmyeong_front_192x160", "jinmyeong_uniform_192x160", "simple_front_192x160"
CHECKED = "2026-11-26 공식 노선정보 대조"

# (번호, 구분, 기점(한), 종점(한), 기점(영), 종점(영), 경유지, 템플릿, 추가 칸, 설명)
BUSAN = [
    ("1001", "급행", "청강리공영차고지", "하단", "Cheonggang-ri Garage", "Hadan",
     "오시리아역,해운대역,센텀시티역,광안역,경성대.부경대역,부산역,사하역\n"
     "Osiria Stn,Haeundae Stn,Centum City Stn,Gwangan Stn,Kyungsung Univ. Stn,Busan Stn,Saha Stn",
     NEW, {"slogan": "Busan is good", "seats": "21"},
     "부산 신형 풀컬러 · 급행. 잔여좌석 21은 화면 예시 값"),
    ("1003", "급행", "기장교리", "부산대학병원", "Gijang Gyo-ri", "PNU Hospital",
     "해운대해수욕장,수영역,광안역,경성대.부경대역,부산역,남포동\n"
     "Haeundae Beach,Suyeong Stn,Gwangan Stn,Kyungsung Univ. Stn,Busan Stn,Nampo-dong",
     OLD, {}, "부산 구형 3색 · 급행 (기점·종점이 길어 두 줄로 나뉨)"),
    ("1004", "급행", "김해", "부산항|국제여객터미널", "Gimhae", "Busan Port Int'l Terminal",
     "대저역,모라,개금,서면역,범일역,부산역\n"
     "Daejeo Stn,Mora,Gaegeum,Seomyeon Stn,Beomil Stn,Busan Stn",
     FC, {}, "풀컬러 방향 · '부산항|국제여객터미널'처럼 | 로 줄 바꿀 자리 지정"),
    ("1002", "심야", "용당공영차고지", "서면", "Yongdang Garage", "Seomyeon",
     "노포역,구서역,부산대학교정문,온천장,사직운동장,부전역\n"
     "Nopo Stn,Guseo Stn,Pusan Nat'l Univ.,Oncheonjang,Sajik Stadium,Bujeon Stn",
     OLD, {}, "1002(심야) · 구분 '심야' 번호 표시"),
    ("86", "일반", "연산9동", "충무동", "Yeonsan 9-dong", "Chungmu-dong",
     "연산역,부산시청,양정역,서면역,범내골역,국제시장,남포동\n"
     "Yeonsan Stn,Busan City Hall,Yangjeong Stn,Seomyeon Stn,Beomnaegol Stn,Gukje Market,Nampo-dong",
     NEW, {"seats": "17"}, "부산 신형 풀컬러 · 일반. 잔여좌석 17은 화면 예시 값"),
    ("40", "일반", "청강리공영차고지", "구덕운동장", "Cheonggang", "Gudeok",
     "송정역,센텀시티역,광안역,경성대.부경대역,부산진역,부산역,국제시장\n"
     "Songjeong Stn,Centum City Stn,Gwangan Stn,Kyungsung Univ. Stn,Busanjin Stn,Busan Stn,"
     "Gukje Market",
     OLD, {}, "부산 구형 3색 · 경유지가 많아 여러 화면으로 넘어감"),
    ("17", "일반", "태종대", "국제백양아파트", "Taejongdae", "Gukje Baegyang APT",
     "남포동,자갈치역,부산역,범내골역,서면역\n"
     "Nampo-dong,Jagalchi Stn,Busan Stn,Beomnaegol Stn,Seomyeon Stn",
     FC, {}, "풀컬러 방향 · 일반"),
    ("중구1", "마을", "충무동교차로", "동대신역", "Chungmu", "Dongdaesin",
     "부평시장,보수사거리,중앙공원,동아대병원,동대신시장\n"
     "Bupyeong Market,Bosu Intersection,Jungang Park,Dong-A Univ. Hospital,Dongdaesin Market",
     OLD, {}, "마을버스 · 번호에 한글이 섞이면 일반 글꼴로 표시"),
]

SEOUL = [
    ("150", "일반", "도봉산역", "시흥대교", "Dobongsan Stn", "Siheung Bridge",
     "쌍문역,수유역,종로2가\n광화문,서울역\n"
     "Ssangmun,Suyu,Jongno 2-ga\n"
     "Gwanghwamun,Seoul Stn",
     JF, {}, "간선 · 진명 전면 192×160"),
    ("360", "일반", "복정역환승센터", "여의도환승센터", "Bokjeong Station", "Yeouido",
     "송파역,잠실역,삼성역\n강남역,논현역,고속터미널\n"
     "Songpa,Jamsil,Samseong\nGangnam,Nonhyeon,Express Term.",
     JF, {}, "간선 · 진명 전면 (기점·종점이 길어 두 줄)"),
    ("470", "일반", "상암차고지", "안골마을", "Sangam Garage", "Angol Village",
     "DMC역,광화문,신사역\n논현역,강남역,양재역\n"
     "DMC,Gwanghwamun,Sinsa\nNonhyeon,Gangnam,Yangjae",
     JU, {}, "간선 · 진명 균일 192×160"),
    ("143", "일반", "정릉", "개포동", "Jeongneung", "Gaepo-dong",
     "삼선교,종로2가\n고속터미널,삼성역,대치역\n"
     "Samseongyo,Jongno 2-ga\nExpress Bus Terminal,Samseong Stn,Daechi Stn",
     SF, {}, "간선 · 단순 전면 192×160"),
    ("7011", "일반", "은평차고지", "중구청", "Eunpyeong Garage", "Jung-gu Office",
     "DMC,합정역,신촌,이대역,서울역,명동역\n"
     "DMC,Hapjeong Stn,Sinchon,Ewha Womans Univ. Stn,Seoul Stn,Myeongdong Stn",
     FC, {}, "지선 · 풀컬러 방향"),
    ("9401", "좌석", "구미동차고지", "서울역", "Gumi-dong Garage", "Seoul Station",
     "정자동,서현역,판교역\n종로,조계사\n"
     "Jeongja-dong,Seohyeon Stn,Pangyo Stn\nJongno,Jogyesa Temple",
     FC, {}, "광역 · 네 자리 번호라 풀컬러 방향 · 구분 '좌석' 색"),
    ("N26", "심야", "강서공영차고지", "중랑공영차고지", "Gangseo Garage", "Jungnang Garage",
     "개화역,마곡역,당산역\n합정역,홍대입구역,신촌\n"
     "Gaehwa,Magok,Dangsan\nHapjeong,Hongdae,Sinchon",
     JF, {}, "심야 · N 번호 (기점·종점은 공식 확인, 경유 정류장은 보조 출처 기준)"),
    ("종로11", "마을", "삼청동", "서울역", "Samcheong-dong", "Seoul Station",
     "삼청공원,국립민속박물관,경복궁,세종문화회관\n"
     "Samcheong Park,Nat'l Folk Museum,Gyeongbokgung,Sejong Center",
     FC, {}, "마을버스 · 풀컬러 방향"),
]

SOURCES = {
    "부산": ["https://bus.busan.go.kr/busanBIMS/ (부산 BIMS 노선검색 · 노선별 정류소 목록, 8개 노선 모두 대조)",
             "https://www.busannavi.net/eng/?p=10173 (1000번대 = 부산 급행버스)",
             "https://info.koreacharts.com/citybus/21/contents.html (처음 작성할 때 쓴 보조 출처)"],
    "서울": ["https://bus.go.kr/ (서울 버스정보 노선 및 정류장 검색 · 150·360·470·143·7011·9401·종로11 정류소 목록 대조, N26은 기점·종점만)",
             "https://info.koreacharts.com/ (N26 경유 정류장 · 처음 작성할 때 쓴 보조 출처)"],
}


def write(path, city, rows):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "노선"
    ws.append([c[1] for c in b.XL_COLS] + ["설명(읽지 않음)"])
    for no, kind, sk, ek, se, ee, via, tpl, extra, note in rows:
        assert tpl in b.all_template_names(), tpl
        d = {"no": no, "kind": kind, "start_ko": sk, "end_ko": ek, "start_en": se, "end_en": ee,
             "via": via, "template": tpl}
        d.update(extra)
        ws.append([d.get(k, "") for k, _ in b.XL_COLS] + [f"{note} · {CHECKED}"])
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFE699")
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for i, w in enumerate([8, 18, 16, 22, 26, 60, 30, 7, 12, 12, 12, 12, 14, 12, 9, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B2"
    g = wb.create_sheet("안내")
    for line in [[f"{city} 실제 노선 샘플 — 노선이미지 생성기"], [""],
                 [f"· 기점·종점·주요 정류장은 {CHECKED} 결과입니다(부산 BIMS · 서울 버스정보시스템). 노선은 바뀔 수 있으니 실제 사용 전 최신 정보로 확인하세요."],
                 ["· 경유지는 주요 정류장만 골랐습니다. 첫 줄(들) = 한글, 다음 줄(들) = 영문."],
                 ["· 잔여좌석·문구 값은 화면 기능을 보여 주려고 넣은 예시입니다."],
                 ["· 프로그램에서 [파일 → 엑셀 불러오기]로 열면 됩니다. 열 제목만 맞으면 순서는 바뀌어도 됩니다."],
                 [""], ["출처"]] + [[s] for s in SOURCES[city]]:
        g.append(line)
    g.column_dimensions["A"].width = 120
    g["A1"].font = Font(bold=True, size=14)
    wb.active = 0
    wb.save(path)


if __name__ == "__main__":
    out = ROOT / "samples"
    write(out / "부산노선_샘플.xlsx", "부산", BUSAN)
    write(out / "서울노선_샘플.xlsx", "서울", SEOUL)
    print("저장:", out / "부산노선_샘플.xlsx", out / "서울노선_샘플.xlsx")
