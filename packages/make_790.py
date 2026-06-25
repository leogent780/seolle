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

def add_schedule_row(table, row_idx, time_text, course_lines, note_text, header_bg=None):
    row = table.rows[row_idx]

    # 시간
    row.cells[0].text = time_text
    for p in row.cells[0].paragraphs:
        for r in p.runs:
            r.font.size = Pt(9)
            r.font.name = 'Malgun Gothic'

    # 코스 (멀티라인)
    cell_text(row.cells[1], course_lines, bold_first=True, font_size=9)

    # 안내
    row.cells[2].text = note_text
    for p in row.cells[2].paragraphs:
        for r in p.runs:
            r.font.size = Pt(9)
            r.font.name = 'Malgun Gothic'

# ─────────────────────────────────────────────
doc = Document()
doc.core_properties.title = 'Real Korean Girl — 1박 2일 패키지'

style = doc.styles['Normal']
style.font.name = 'Malgun Gothic'
style.font.size = Pt(10)

# 제목
add_heading(doc, '💅 Real Korean Girl — 1박 2일 K뷰티 스타일링 패키지', 1, (31, 73, 125))
p = doc.add_paragraph('현지 한국 여성처럼, 자연스럽고 세련된 데일리 K뷰티 변신')
p.runs[0].italic = True
p.runs[0].font.name = 'Malgun Gothic'
p2 = doc.add_paragraph('명동·강남 일대  |  큐레이터 동행 옵션')
p2.runs[0].italic = True
p2.runs[0].font.name = 'Malgun Gothic'

# ── DAY 1 ──
doc.add_paragraph()
add_heading(doc, 'DAY 1', 2, (31, 73, 125))

day1_data = [
    (
        '10:00\n~\n11:30',
        [
            '퍼스널컬러 분석 1:1  (90분)',
            '· 웜톤/쿨톤 세부 계절 유형 분류 (봄·여름·가을·겨울)',
            '· 얼굴형 & 체형 분석 / 통역 포함',
            '',
            '추천 업체 TOP5',
            '① 로아 퍼스널컬러  ② 퍼스널컬러랩  ③ 뮤즈클로젯',
            '④ 위드컬러  ⑤ 마이컬러스튜디오',
            '',
            '💡 오늘 쇼핑 방향과 내일 메이크업 컬러를 결정하는 핵심 첫 단계.',
            '   이 결과 없이는 이후 모든 코스가 흐릿해짐',
        ],
        '✅ 포함'
    ),
    (
        '11:30\n~\n13:30',
        [
            '올리브영 & 약국 쇼핑  (2시간)',
            '· 퍼스널컬러 결과 직후 바로 쇼핑 적용',
            '· 명동 올리브영 메가스토어',
            '· 명동 인근 약국 (Dr.G, 미샤, 이니스프리)',
            '',
            '💡 진단 결과가 가장 선명할 때 바로 쇼핑.',
            '   시간이 지날수록 기억이 희미해짐',
        ],
        '💳 쇼핑비\n개별 결제'
    ),
    (
        '13:30\n~\n14:30',
        [
            '설화수 플래그십 팝업 체험  (1시간)',
            '· 한방 뷰티 브랜드 프리미엄 체험',
            '· 전용 스킨케어 컨설팅',
            '· SNS 포토존 촬영',
            '',
            '💡 쇼핑 후 자연스러운 동선.',
            '   한국 대표 럭셔리 뷰티 브랜드 경험으로 패키지 고급감 UP',
        ],
        '✅ 포함'
    ),
    (
        '14:30\n~\n15:30',
        [
            '티 오마카세  (1시간)',
            '· 전통차 소믈리에가 선별한 계절 차 코스 제공',
            '· 차 설명 + 다과 페어링 포함',
            '· 명동·인사동 인근 프리미엄 티룸',
            '',
            '추천 업체 TOP5',
            '① 오설록 티하우스  ② 차마시는 뜰  ③ 티컬렉티브',
            '④ 명가원  ⑤ 오우재',
            '',
            '💡 에스테틱 전 충분한 휴식.',
            '   프리미엄 차 코스로 패키지 고급감 UP — 긴장 풀린 상태에서 시술 효과 극대화',
        ],
        '✅ 포함'
    ),
    (
        '15:30\n~\n17:30',
        [
            '피부진단 & 에스테틱  (2시간)',
            '· 전문 에스테티션 1:1 피부 진단 (20분)',
            '· 시술 선택 (60–80분)',
            '  — 트러블집중관리 + 진정관리',
            '  — 얼굴윤곽 + 탄력관리',
            '  — 화이트닝 + 탄력관리',
            '',
            '추천 업체 TOP5',
            '① 오슬 스킨케어  ② 라뷰 에스테틱  ③ 더스킨랩',
            '④ 뷰티풀스킨  ⑤ 스킨케어 클리닉',
            '',
            '💡 헤어&메이크업 전날 미리 피부 정돈.',
            '   하룻밤 자고 나면 피부가 안정되어 다음날 화장이 훨씬 잘 먹음',
        ],
        '✅ 포함'
    ),
    (
        '19:30',
        ['DAY 1 해산', '· 명동역 인근'],
        '—'
    ),
]

table1 = doc.add_table(rows=1 + len(day1_data), cols=3)
table1.style = 'Table Grid'
for i, h in enumerate(['시간', '코스', '안내']):
    cell = table1.rows[0].cells[i]
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10)
    cell.paragraphs[0].runs[0].font.name = 'Malgun Gothic'
    set_cell_bg(cell, 'D9E2F3')

for i, (t, c, n) in enumerate(day1_data):
    add_schedule_row(table1, i + 1, t, c, n)

# 열 너비 조정
for row in table1.rows:
    row.cells[0].width = Cm(2.0)
    row.cells[1].width = Cm(11.5)
    row.cells[2].width = Cm(2.5)

# ── DAY 2 ──
doc.add_paragraph()
add_heading(doc, 'DAY 2', 2, (31, 73, 125))

day2_data = [
    (
        '10:00\n~\n12:00',
        [
            '데일리 헤어 & 메이크업  (2시간)',
            '· 한국식 데일리 헤어 스타일링',
            '· 퍼스널컬러 기반 내추럴 글로우 메이크업',
            '· 전날 에스테틱 피부 상태 고려한 베이스 메이크업',
            '',
            '추천 업체 TOP5',
            '① 박준뷰티랩  ② 준오헤어  ③ 이철헤어커커',
            '④ 토니앤가이  ⑤ 헤라뷰티스튜디오',
            '',
            '💡 전날 에스테틱으로 피부가 최상 컨디션.',
            '   이 타이밍에 메이크업해야 화장이 가장 잘 받고 오래 유지됨',
        ],
        '✅ 포함'
    ),
    (
        '12:00\n~\n14:00',
        [
            '프로필 촬영  (2시간)',
            '· 명동 인근 실내 스튜디오',
            '· 전문 포토그래퍼',
            '· 자연스러운 데일리 룩 + 글로우 컨셉',
            '',
            '추천 업체 TOP5',
            '① 필름드 스튜디오  ② 모멘트 포토  ③ 셀프리 스튜디오',
            '④ 데일리 스냅  ⑤ 피치 스튜디오',
            '',
            '💡 모든 준비가 완성된 최상의 상태.',
            '   헤어·메이크업 직후 바로 촬영해야 가장 완벽한 컷이 나옴',
        ],
        '✅ 포함'
    ),
    (
        '14:00',
        ['해산', '· 명동역 인근'],
        '—'
    ),
    (
        '별도\n날짜',
        [
            '명동 TOP5 병원 시술 예약 대행',
            '· 퍼스널컬러 진단 결과 기반 시술 추천',
            '· 피부과 / 성형외과 중 선택',
            '· 시술비는 병원 현장 결제',
        ],
        '🏥 예약\n대행'
    ),
]

table2 = doc.add_table(rows=1 + len(day2_data), cols=3)
table2.style = 'Table Grid'
for i, h in enumerate(['시간', '코스', '안내']):
    cell = table2.rows[0].cells[i]
    cell.text = h
    cell.paragraphs[0].runs[0].bold = True
    cell.paragraphs[0].runs[0].font.size = Pt(10)
    cell.paragraphs[0].runs[0].font.name = 'Malgun Gothic'
    set_cell_bg(cell, 'D9E2F3')

for i, (t, c, n) in enumerate(day2_data):
    add_schedule_row(table2, i + 1, t, c, n)

for row in table2.rows:
    row.cells[0].width = Cm(2.0)
    row.cells[1].width = Cm(11.5)
    row.cells[2].width = Cm(2.5)

# ── 포함 항목 ──
doc.add_paragraph()
add_heading(doc, '🎁 포함 항목', 2)
add_heading(doc, 'Real Korean Girl 패키지 포함', 3)
for item in [
    '퍼스널컬러 분석 1:1 (통역포함)',
    '설화수 팝업 체험',
    '티 오마카세 (전통차 코스 + 다과)',
    '피부진단 & 에스테틱',
    '데일리 헤어 & 메이크업',
    '프로필 촬영',
    '웰컴 K뷰티 샘플팩',
]:
    p = doc.add_paragraph(item, style='List Bullet')
    p.runs[0].font.name = 'Malgun Gothic'

add_heading(doc, '여행 편의 키트', 3)
for item in [
    '한국 eSIM',
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
p = doc.add_paragraph()
r = p.add_run('기본 패키지  ₩790,000')
r.bold = True
r.font.size = Pt(12)
r.font.name = 'Malgun Gothic'

doc.add_paragraph()
add_simple_table(doc,
    ['큐레이터 동행 구간', '추가 요금', '총 금액'],
    [
        ['없음', '—', '₩790,000'],
        ['1구간', '+₩50,000', '₩840,000'],
        ['2구간', '+₩80,000', '₩870,000'],
        ['3구간', '+₩110,000', '₩900,000'],
        ['4구간', '+₩130,000', '₩920,000'],
        ['5구간 (전 구간)', '+₩150,000', '₩940,000'],
    ],
    'D9E2F3'
)
doc.add_paragraph('※ 큐레이터는 원하는 구간만 연속으로 선택 가능합니다').runs[0].font.name = 'Malgun Gothic'
doc.add_paragraph('※ 지원 언어: 영어 · 중국어 · 일본어').runs[0].font.name = 'Malgun Gothic'

doc.add_paragraph()
add_simple_table(doc,
    ['항목', '결제'],
    [
        ['명동 병원 시술비', '💳 현장결제 (별도 날짜)'],
        ['쇼핑비', '💳 개별 결제'],
    ],
    'FFF2CC'
)

# ── 원가 구조 (내부) ──
doc.add_paragraph()
add_heading(doc, '📊 원가 구조 (내부 전용)', 2, (192, 0, 0))
add_simple_table(doc,
    ['항목', '원가'],
    [
        ['퍼스널컬러 분석 1:1 (통역포함)', '₩148,000'],
        ['피부진단 & 에스테틱', '₩80,000'],
        ['데일리 헤어 & 메이크업', '₩154,000'],
        ['프로필 촬영', '₩100,000'],
        ['티 오마카세', '₩40,000'],
        ['웰컴 K뷰티 샘플팩', '₩8,000'],
        ['번들 (eSIM·보험 등)', '₩12,000'],
        ['운영비', '₩20,000'],
        ['원가 합계', '₩562,000'],
        ['원가율', '71.1% ⚠️'],
    ],
    'FCE4D6'
)

out = '/home/user/seolle/packages/고객공유용_79만원_RealKoreanGirl.docx'
doc.save(out)
print('완료:', out)
