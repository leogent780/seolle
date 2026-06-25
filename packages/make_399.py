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
doc.core_properties.title = 'Seoul Prestige — 2박 3일 울트라 럭셔리 뷰티'

style = doc.styles['Normal']
style.font.name = 'Malgun Gothic'
style.font.size = Pt(10)

add_heading(doc, '👑 Seoul Prestige — 2박 3일 울트라 럭셔리 뷰티', 1, (31, 73, 125))
p = doc.add_paragraph('연예인 담당 아티스트·화보급 촬영·의원급 시술로 완성하는 최상위 K뷰티 경험')
p.runs[0].italic = True
p.runs[0].font.name = 'Malgun Gothic'
p2 = doc.add_paragraph('강남·경복궁 일대  |  전담 큐레이터 3일 풀 동행')
p2.runs[0].italic = True
p2.runs[0].font.name = 'Malgun Gothic'

# ── 비교표 ──
doc.add_paragraph()
add_heading(doc, '📊 ₩390,000 vs ₩3,990,000 비교', 2)
add_simple_table(doc,
    ['항목', '₩390,000 기본', '₩3,990,000 프레스티지'],
    [
        ['퍼스널컬러', '일반 살롱 (그룹)', '연예인 담당 컬러리스트 1:1 + 골격진단'],
        ['헤어&메이크업', '기본 글로우', '연예인 담당 아티스트 (2.5시간 화보급)'],
        ['프로필 촬영', '기본 스튜디오 1컨셉', '화보급 스튜디오 + 야외 로케이션 3컨셉 · 보정본 30장+'],
        ['큐레이터', '없음', '전담 큐레이터 3일 풀 동행'],
        ['에스테틱', '없음', '의원급 피부과 연계 시술'],
        ['반영구', '없음', '연예인 담당 1:1 아티스트'],
        ['디너', '없음', '한식 오마카세 코스'],
        ['기간', '당일', '2박 3일'],
    ],
    'E2EFDA'
)

# ── DAY 1 ──
doc.add_paragraph()
add_heading(doc, 'DAY 1 — 럭셔리 진단 & 피부', 2, (31, 73, 125))

day1_data = [
    (
        '10:00\n~\n12:00',
        [
            '퍼스널컬러 & 골격진단 1:1  (2시간)',
            '· 연예인·방송 담당 컬러리스트 전담 배정',
            '· 웜톤/쿨톤 세부 계절 유형 + 퍼스널컬러 팔레트 제작',
            '· 얼굴형·체형·골격 심층 분석',
            '· 통역 포함',
            '',
            '추천 업체 TOP5',
            '① 로아 퍼스널컬러  ② 퍼스널컬러랩  ③ 뮤즈클로젯',
            '④ 위드컬러  ⑤ 마이컬러스튜디오',
            '',
            '💡 39만원 패키지와 같은 진단이지만 연예인 담당 컬러리스트가 2배 깊이로 분석.',
            '   이 결과가 3일 전체 시술·스타일링의 기준이 됨',
        ],
        '✅ 포함'
    ),
    (
        '13:00\n~\n15:00',
        [
            '에스테틱 (의원급)  (2시간)',
            '· 의원 연계 피부과 전문의 피부 진단',
            '· 프리미엄 피부 시술 1종 선택',
            '  — 리쥬란힐러 + 물광주사',
            '  — 울쎄라 리프팅',
            '  — 레이저토닝 + 수분광채',
            '',
            '추천 업체 TOP5',
            '① 강남 유명 피부과  ② 청담 스킨클리닉  ③ 압구정 피부과',
            '④ 명동 더마클리닉  ⑤ 신사동 에스테틱 메디컬',
            '',
            '💡 일반 에스테틱이 아닌 의원급 시술 — 결과가 다음 날부터 바로 피부에 나타남',
        ],
        '✅ 포함'
    ),
    (
        '15:30\n~\n17:30',
        [
            '명품 뷰티 쇼핑 동행  (2시간)',
            '· 신세계 강남 or 갤러리아 백화점 뷰티 플로어',
            '· La Mer · Sulwhasoo · SK-II · 헤라 · 로레알 럭셔리',
            '· 퍼스널컬러 기반 1:1 제품 큐레이션',
            '· 면세 & 멤버십 혜택 가이드',
            '',
            '💡 진단 당일 최고급 브랜드에서 내 컬러 맞춤 제품 구매.',
            '   일반 올리브영 쇼핑과의 가장 큰 차별점',
        ],
        '💳 쇼핑비\n개별 결제'
    ),
    (
        '17:30\n~\n19:00',
        [
            '네일 케어 럭셔리  (90분)',
            '· 프리미엄 젤 네일 (퍼스널컬러 기반 컬러 추천)',
            '· 핸드 케어 + 큐티클 정리 + 파라핀 트리트먼트',
            '',
            '추천 업체 TOP5',
            '① 네일팝  ② 에뛰드네일  ③ 글로시네일',
            '④ 핑크네일명동  ⑤ 네일아트클럽',
            '',
            '💡 쇼핑 직후 퍼스널컬러 기반 네일 컬러 선택.',
            '   여행 첫날 손끝까지 완성하면 남은 2일이 더 빛남',
        ],
        '✅ 포함'
    ),
    (
        '19:30\n~\n21:30',
        [
            '한식 오마카세 디너  (2시간)',
            '· 경복궁 인근 프리미엄 한정식 오마카세',
            '· 큐레이터 동행 & 메뉴 설명',
            '',
            '추천 업체 TOP5',
            '① 한식공간  ② 지화자  ③ 밍글스',
            '④ 라궁  ⑤ 온지음',
            '',
            '💡 하루의 피로를 풀며 내일 시술을 위한 영양 보충.',
            '   코스 한식으로 한국 음식 문화까지 경험',
        ],
        '✅ 포함'
    ),
    (
        '22:00',
        ['DAY 1 해산'],
        '—'
    ),
]

add_day_table(doc, day1_data)

# ── DAY 2 ──
doc.add_paragraph()
add_heading(doc, 'DAY 2 — 반영구 & 스페셜 케어', 2, (31, 73, 125))

day2_data = [
    (
        '10:00\n~\n12:30',
        [
            '반영구 시술 1종 선택  (2.5시간)',
            '· 전담 1:1 아티스트 배정 (연예인·유튜버 포트폴리오 보유)',
            '· 눈썹 문신 — 자연 결 눈썹',
            '· 립 문신 — 혈색 핑크/코럴',
            '· 아이라인 문신 — 또렷한 눈매',
            '· 두피 문신 — 헤어라인·볼륨 보정',
            '',
            '추천 업체 TOP5',
            '① 오브뷰티  ② 미소반영구  ③ 아트메이크업랩',
            '④ 뷰티펙트  ⑤ 라인아티스트',
            '',
            '💡 39만원 패키지에 없는 반영구 — 한 번 시술로 1~2년 유지.',
            '   내일 촬영 전 얼굴 전체 밸런스를 완성하는 핵심 단계',
        ],
        '✅ 포함'
    ),
    (
        '13:30\n~\n15:00',
        [
            '헤드스파 or 전신 마사지 선택  (90분)',
            '· 헤드스파 — 두피 진단 + 스케일링 + 한방 트리트먼트',
            '· 전신 마사지 — 아로마 오일 전신, 시술 피로 회복',
            '',
            '추천 업체 TOP5',
            '① 모스트헤어  ② 헤드스파클리닉  ③ 힐링스파명동',
            '④ 바디웍스  ⑤ 아로마힐링강남',
            '',
            '💡 반영구 직후 긴장된 몸을 풀어주는 리커버리 타임',
        ],
        '✅ 포함'
    ),
    (
        '15:30\n~\n16:30',
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
            '💡 반영구·마사지 후 차분하게 마무리.',
            '   내일 촬영을 앞두고 몸과 마음을 고요하게 정돈',
        ],
        '✅ 포함'
    ),
    (
        '17:00',
        ['DAY 2 해산'],
        '—'
    ),
]

add_day_table(doc, day2_data)

# ── DAY 3 ──
doc.add_paragraph()
add_heading(doc, 'DAY 3 — 화보 촬영', 2, (31, 73, 125))

day3_data = [
    (
        '10:00\n~\n12:30',
        [
            '헤어 & 메이크업 (연예인 담당 아티스트)  (2.5시간)',
            '· 드라마·광고·화보 담당 아티스트 전담 배정',
            '· 퍼스널컬러·골격·반영구 결과 통합 적용',
            '· 3컨셉 대응 멀티 스타일링',
            '',
            '추천 업체 TOP5',
            '① 박준뷰티랩 VIP  ② 헤라뷰티스튜디오  ③ 준오헤어 프레스티지',
            '④ 이철헤어커커  ⑤ 토니앤가이 시그니처',
            '',
            '💡 39만원 H&M과의 차이: 연예인 담당 아티스트 전담 + 2.5시간 풀 스타일링.',
            '   메이크업이 달라지면 촬영 결과물이 완전히 달라짐',
        ],
        '✅ 포함'
    ),
    (
        '13:00\n~\n17:00',
        [
            '화보급 프로필 촬영  (4시간)',
            '· 프리미엄 스튜디오 (강남/청담)',
            '· 드라마·광고 담당 포토그래퍼',
            '· 컨셉 1 — 현대 글로우 룩 (스튜디오)',
            '· 컨셉 2 — 한복 전통 컨셉 (스튜디오)',
            '· 컨셉 3 — 야외 로케이션 (경복궁 or 청담)',
            '· 보정본 30장 이상 제공',
            '',
            '추천 업체 TOP5',
            '① 필름드스튜디오 프리미엄  ② 모멘트포토  ③ 청담 스튜디오K',
            '④ 데일리스냅 VIP  ⑤ 피치스튜디오',
            '',
            '💡 39만원 기본 스튜디오(1컨셉)와의 차이: 3컨셉 + 야외 로케이션 + 드라마급 포토그래퍼.',
            '   보정본 30장 이상 = 진짜 화보',
        ],
        '✅ 포함'
    ),
    (
        '17:30',
        ['해산'],
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
add_heading(doc, 'Seoul Prestige 패키지 포함', 3)
for item in [
    'DAY1  퍼스널컬러 & 골격진단 1:1 (연예인 담당 컬러리스트, 통역포함)',
    'DAY1  의원급 에스테틱 시술 1종',
    'DAY1  럭셔리 네일 케어',
    'DAY1  한식 오마카세 코스 디너',
    'DAY2  반영구 시술 1종 (연예인 담당 1:1 아티스트)',
    'DAY2  헤드스파 or 전신 마사지 (선택)',
    'DAY3  헤어 & 메이크업 (연예인 담당 아티스트)',
    'DAY3  화보급 프로필 촬영 3컨셉 + 보정본 30장+',
    '전담 큐레이터 3일 풀 동행 (통역·쇼핑 동행·병원 동행)',
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
add_simple_table(doc,
    ['구분', '금액'],
    [
        ['Seoul Prestige 2박 3일', '₩3,990,000'],
        ['명동 병원 시술비', '💳 현장결제 (별도 날짜)'],
        ['명품 뷰티 쇼핑비', '💳 개별 결제'],
    ],
    'D9E2F3'
)

# ── 원가 구조 (내부) ──
doc.add_paragraph()
add_heading(doc, '📊 원가 구조 (내부 전용)', 2, (192, 0, 0))
add_simple_table(doc,
    ['항목', '원가'],
    [
        ['퍼스널컬러 & 골격진단 (연예인급)', '₩300,000'],
        ['에스테틱 (의원급)', '₩150,000'],
        ['한식 오마카세 디너', '₩120,000'],
        ['반영구 시술 1종 (최고급)', '₩200,000'],
        ['헤드스파 or 전신 마사지', '₩80,000'],
        ['네일 케어 (럭셔리)', '₩50,000'],
        ['헤어 & 메이크업 (연예인 담당)', '₩350,000'],
        ['화보급 스튜디오 + 로케이션 + 한복 컨셉', '₩500,000'],
        ['전담 큐레이터 3일 (지급 80%)', '₩480,000'],
        ['번들 (eSIM·보험 등)', '₩25,000'],
        ['운영비', '₩80,000'],
        ['원가 합계', '₩2,335,000'],
        ['원가율', '58.5% ✅'],
    ],
    'FCE4D6'
)

p = doc.add_paragraph()
r = p.add_run('※ 촬영 원가(₩780,000)는 스튜디오·로케이션·한복 컨셉 통합 기준 예상가. 업체 협약 후 조정 가능.')
r.font.size = Pt(9)
r.font.name = 'Malgun Gothic'
r.italic = True

out = '/home/user/seolle/packages/고객공유용_399만원_SeoulPrestige.docx'
doc.save(out)
print('완료:', out)
