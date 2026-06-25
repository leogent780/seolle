from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def cell_text(cell, lines, bold_first=False, font_size=9):
    cell.paragraphs[0].clear()
    for i, line in enumerate(lines):
        if i == 0:
            p = cell.paragraphs[0]
        else:
            p = cell.add_paragraph()
        run = p.add_run(line)
        run.font.size = Pt(font_size)
        run.font.name = 'Malgun Gothic'
        if i == 0 and bold_first:
            run.bold = True
        p.paragraph_format.space_before = Pt(1)
        p.paragraph_format.space_after = Pt(1)

def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        run.font.name = 'Malgun Gothic'
        if color:
            run.font.color.rgb = RGBColor(*color)
    return h

def add_simple_table(doc, headers, rows, header_bg='D9E2F3'):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].font.name = 'Malgun Gothic'
        set_cell_bg(cell, header_bg)
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            row.cells[c_idx].text = str(val)
            row.cells[c_idx].paragraphs[0].runs[0].font.size = Pt(10)
            row.cells[c_idx].paragraphs[0].runs[0].font.name = 'Malgun Gothic'
    return table

def add_schedule_row(table, row_idx, time_text, course_lines, note_text):
    row = table.rows[row_idx]
    row.cells[0].text = time_text
    for p in row.cells[0].paragraphs:
        for r in p.runs:
            r.font.size = Pt(9)
            r.font.name = 'Malgun Gothic'
    cell_text(row.cells[1], course_lines, bold_first=True, font_size=9)
    row.cells[2].text = note_text
    for p in row.cells[2].paragraphs:
        for r in p.runs:
            r.font.size = Pt(9)
            r.font.name = 'Malgun Gothic'

def add_day_table(doc, day_data):
    table = doc.add_table(rows=1 + len(day_data), cols=3)
    table.style = 'Table Grid'
    for i, h in enumerate(['시간', '코스', '안내']):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        cell.paragraphs[0].runs[0].font.name = 'Malgun Gothic'
        set_cell_bg(cell, 'D9E2F3')
    for i, (t, c, n) in enumerate(day_data):
        add_schedule_row(table, i + 1, t, c, n)
    for row in table.rows:
        row.cells[0].width = Cm(2.0)
        row.cells[1].width = Cm(11.5)
        row.cells[2].width = Cm(2.5)
    return table

# ─────────────────────────────────────────────
doc = Document()
doc.core_properties.title = 'Seoul Traditional Beauty — 2박 3일 전통체험 뷰티 패키지'

style = doc.styles['Normal']
style.font.name = 'Malgun Gothic'
style.font.size = Pt(10)

add_heading(doc, '🏯 Seoul Traditional Beauty — 2박 3일 전통체험 뷰티 패키지', 1, (31, 73, 125))
p = doc.add_paragraph('한국의 전통미와 현대 K뷰티를 함께 경험하는 문화 여정')
p.runs[0].italic = True
p.runs[0].font.name = 'Malgun Gothic'
p2 = doc.add_paragraph('강남·경복궁 일대  |  전담 큐레이터 동행')
p2.runs[0].italic = True
p2.runs[0].font.name = 'Malgun Gothic'

# ── DAY 1 ──
doc.add_paragraph()
add_heading(doc, 'DAY 1 — 전통 진단 & 뷰티', 2, (31, 73, 125))

day1_data = [
    (
        '10:00\n~\n11:30',
        [
            '퍼스널컬러 & 골격진단 1:1',
            '· 웜톤/쿨톤 세부 계절 유형 분류 (봄·여름·가을·겨울)',
            '· 얼굴형 & 체형 골격진단',
            '· 통역 포함',
            '',
            '추천 업체 TOP5',
            '① 로아 퍼스널컬러  ② 퍼스널컬러랩  ③ 뮤즈클로젯',
            '④ 위드컬러  ⑤ 마이컬러스튜디오',
            '',
            '💡 3일 전체 스타일링·시술의 방향을 결정하는 핵심 첫 단계',
        ],
        '✅ 포함'
    ),
    (
        '12:00\n~\n14:00',
        [
            '에스테틱 (1회 선택)',
            '· 트러블·진정·윤곽·화이트닝 관리',
            '· 물광아쿠아필 + 쿨링이온',
            '· 모공 타이트닝 + 보습 관리',
            '',
            '추천 업체 TOP5',
            '① 오슬 스킨케어  ② 라뷰 에스테틱  ③ 더스킨랩',
            '④ 뷰티풀스킨  ⑤ 스킨케어 클리닉',
        ],
        '✅ 포함'
    ),
    (
        '14:30\n~\n17:30',
        [
            '한복 대여 & 경복궁·북촌 포토워크',
            '· 프리미엄 한복 선택 착용',
            '· 경복궁·북촌 한옥마을 큐레이터 동행',
            '· 포토 스팟 가이드',
            '',
            '추천 업체 TOP5',
            '① 리(Re)한복  ② 황후의날개  ③ 스튜디오 어울림',
            '④ 한복남  ⑤ 궁중한복',
            '',
            '💡 퍼스널컬러 진단 직후 — 어울리는 색 한복 바로 적용',
        ],
        '✅ 포함'
    ),
    (
        '18:00\n~\n20:00',
        [
            '전통 한식 코스 디너',
            '· 경복궁 인근 프리미엄 한정식',
            '· 큐레이터 동행',
            '',
            '추천 업체 TOP5',
            '① 한식공간  ② 지화자  ③ 밍글스',
            '④ 라궁  ⑤ 온지음',
            '',
            '💡 전통 음식 문화 체험으로 한국 여행의 완성',
        ],
        '✅ 포함'
    ),
    (
        '20:00',
        ['DAY 1 해산', '· 경복궁역 인근'],
        '—'
    ),
]

add_day_table(doc, day1_data)

# ── DAY 2 ──
doc.add_paragraph()
add_heading(doc, 'DAY 2 — 전통 문화 & 힐링 체험', 2, (31, 73, 125))

day2_data = [
    (
        '10:00\n~\n11:30',
        [
            '다도 체험 + 사주',
            '· 전통 다구 사용법 + 명상 다도 (1시간)',
            '· 전통 사주 분석 (통역 포함, 30분)',
            '',
            '추천 업체 TOP5',
            '① 인사동 다도원  ② 차마시는 뜰  ③ 오설록 티하우스',
            '④ 전통찻집 귀천  ⑤ 명가원',
            '',
            '💡 하루를 여는 명상 다도 — 전통 문화의 깊이를 먼저 경험',
        ],
        '✅ 포함'
    ),
    (
        '12:00\n~\n13:30',
        [
            '효소찜질 or 황토팩 or 1인 세신 (선택 1)',
            '· 효소찜질 — 발효 효소 찜질, 체온 상승·노폐물 배출',
            '· 황토팩 전신케어 — 황토·숯 전신팩, 피부 해독·미백',
            '· 1인 세신 — 전통 이탈리아 타월 전신 세신, 각질 제거',
            '',
            '추천 업체 TOP5',
            '① 황토사우나 본점  ② 자연황토체험관  ③ 힐링황토방',
            '④ 효소찜질원  ⑤ 전통황토스파',
            '',
            '💡 취향에 맞게 몸 안팎 정화 — 내일 촬영 전 최상의 피부 준비',
        ],
        '✅ 포함'
    ),
    (
        '14:00\n~\n15:30',
        [
            '막걸리 빚기 or 전통주 체험 (선택 1)',
            '· 전통 발효주 직접 제조 과정 체험',
            '· 시음 포함',
            '',
            '추천 업체 TOP5',
            '① 서울막걸리 체험장  ② 배상면주가  ③ 국순당 양조장',
            '④ 전통주갤러리  ⑤ 막걸리학교',
            '',
            '💡 한국 전통 발효 문화를 몸으로 배우는 시간',
        ],
        '✅ 포함'
    ),
    (
        '16:00\n~\n18:00',
        [
            '당일 템플스테이 (은평 한옥마을)  (4시간)',
            '· 09:30  접수 및 오리엔테이션',
            '· 10:00  사찰 안내 및 역사 소개',
            '· 10:30  참선 (명상)',
            '· 11:10  스님과 차담',
            '· 11:40  사찰 음식 (공양)',
            '· 12:30  자유 산책 및 회향',
            '',
            '추천 업체 TOP5',
            '① 은평 진관사  ② 북한산 삼천사  ③ 봉은사',
            '④ 길상사  ⑤ 조계사',
            '',
            '💡 바쁜 여행 일정 속 진짜 쉬어가는 시간.',
            '   참선·차담·사찰음식으로 한국 불교 문화를 온몸으로 경험',
        ],
        '✅ 포함'
    ),
    (
        '19:30',
        ['DAY 2 해산', '· 명동역 인근'],
        '—'
    ),
]

add_day_table(doc, day2_data)

# ── DAY 3 ──
doc.add_paragraph()
add_heading(doc, 'DAY 3 — 피니싱 & 촬영', 2, (31, 73, 125))

day3_data = [
    (
        '10:00\n~\n12:00',
        [
            '올리브영 & 약국 쇼핑',
            '· 퍼스널컬러 기반 맞춤 제품 큐레이션',
            '· 면세 쇼핑 가이드',
        ],
        '💳 쇼핑비\n개별 결제'
    ),
    (
        '13:00\n~\n15:00',
        [
            '헤어 & 메이크업 (연주회/시상식급)',
            '· 골격·퍼스널컬러 기반 맞춤 스타일링',
            '',
            '추천 업체 TOP5',
            '① 박준뷰티랩  ② 준오헤어  ③ 이철헤어커커',
            '④ 토니앤가이  ⑤ 헤라뷰티스튜디오',
            '',
            '💡 DAY2 한방 케어로 피부 최상 컨디션 — 화장이 가장 잘 받는 타이밍',
        ],
        '✅ 포함'
    ),
    (
        '15:00\n~\n17:30',
        [
            '프리미엄 프로필 촬영',
            '· 프리미엄 스튜디오',
            '· 전문 포토그래퍼',
            '· 한복 + 현대 의상 멀티 컨셉',
            '',
            '추천 업체 TOP5',
            '① 필름드 스튜디오  ② 모멘트 포토  ③ 셀프리 스튜디오',
            '④ 데일리 스냅  ⑤ 피치 스튜디오',
            '',
            '💡 3일간의 변화가 모두 담긴 최상의 상태에서 촬영',
        ],
        '✅ 포함'
    ),
    (
        '17:30',
        ['해산', '· 명동역 인근'],
        '—'
    ),
    (
        '별도\n날짜',
        [
            '명동 TOP5 병원 시술 예약 대행',
            '· 퍼스널컬러·골격 진단 결과 기반 시술 추천',
            '· 피부과 / 성형외과 중 선택',
            '· 큐레이터 병원 동행 & 통역',
            '· 시술비는 병원 현장 결제',
        ],
        '🏥 예약\n대행'
    ),
]

add_day_table(doc, day3_data)

# ── 포함 항목 ──
doc.add_paragraph()
add_heading(doc, '🎁 포함 항목', 2)
add_heading(doc, '전통체험 뷰티 패키지 포함', 3)
for item in [
    'DAY1  퍼스널컬러 & 골격진단 1:1 (통역포함)',
    'DAY1  에스테틱',
    'DAY1  프리미엄 한복 대여',
    'DAY1  경복궁·북촌 큐레이터 동행 포토워크',
    'DAY1  전통 한식 코스 디너',
    'DAY2  다도 체험 + 사주',
    'DAY2  효소찜질 or 황토팩 or 1인 세신 (선택)',
    'DAY2  막걸리 빚기 or 전통주 체험 (선택)',
    'DAY2  당일 템플스테이 (은평 한옥마을)',
    'DAY3  헤어 & 메이크업 (연주회/시상식급)',
    'DAY3  프리미엄 프로필 촬영 (멀티 컨셉)',
    '전담 큐레이터 3일 풀 동행',
]:
    p = doc.add_paragraph(item, style='List Bullet')
    p.runs[0].font.name = 'Malgun Gothic'

add_heading(doc, '여행 편의 키트', 3)
for item in [
    '한국 eSIM',
    'T-money 교통카드',
    '여행자 보험',
    '코스별 예약 확인서',
    '카카오톡 긴급 연락',
    '면세 쇼핑 쿠폰',
]:
    p = doc.add_paragraph(item, style='List Bullet')
    p.runs[0].font.name = 'Malgun Gothic'

# ── 가격 안내 ──
doc.add_paragraph()
add_heading(doc, '💰 가격 안내', 2)
add_simple_table(doc,
    ['구분', '금액'],
    [
        ['전통체험 뷰티 2박 3일', '₩2,500,000'],
        ['명동 병원 시술비', '💳 현장결제 (별도 날짜)'],
        ['쇼핑비', '💳 개별 결제'],
    ],
    'D9E2F3'
)

# ── 원가 구조 (내부) ──
doc.add_paragraph()
add_heading(doc, '📊 원가 구조 (내부 전용)', 2, (192, 0, 0))
add_simple_table(doc,
    ['항목', '원가'],
    [
        ['퍼스널컬러 & 골격진단 1:1', '₩260,000'],
        ['에스테틱', '₩80,000'],
        ['한복 대여 (프리미엄)', '₩40,000'],
        ['전통 한식 코스 디너', '₩70,000'],
        ['다도 체험 + 사주', '₩50,000'],
        ['효소찜질 or 황토팩 or 1인 세신', '₩60,000'],
        ['막걸리/전통주 체험', '₩40,000'],
        ['당일 템플스테이 (은평 한옥마을)', '₩50,000'],
        ['헤어 & 메이크업 (연주회급)', '₩150,000'],
        ['프리미엄 프로필 촬영', '₩270,000'],
        ['전담 큐레이터 3일 (지급 80%)', '₩480,000'],
        ['번들 (eSIM·T-money·보험 등)', '₩25,000'],
        ['운영비', '₩50,000'],
        ['원가 합계', '₩1,625,000'],
        ['원가율', '65.0% ⚠️'],
    ],
    'FCE4D6'
)

out = '/home/user/seolle/packages/고객공유용_250만원_SeoulTraditionalBeauty.docx'
doc.save(out)
print('완료:', out)
