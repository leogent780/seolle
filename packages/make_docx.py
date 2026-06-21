from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for run in h.runs:
        if color:
            run.font.color.rgb = RGBColor(*color)
    return h

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_table(doc, headers, rows, header_bg=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = h
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(10)
        if header_bg:
            set_cell_bg(cell, header_bg)
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(val)
            cell.paragraphs[0].runs[0].font.size = Pt(10)
    return table

def add_schedule_table(doc, header_bg='D9E2F3'):
    """하루 일정 표 — 코스 셀에 세부 항목을 줄바꿈으로 추가"""
    schedule = [
        {
            'time': '10:00',
            'title': '퍼스널컬러 분석',
            'details': [
                '웜톤/쿨톤 세부 계절 유형 분류 (봄·여름·가을·겨울)',
                '얼굴형 분석',
                '체형 분석',
            ],
            'note': '✅ 포함',
        },
        {
            'time': '11:00',
            'title': '올리브영 & 약국 쇼핑',
            'details': [
                '퍼스널컬러 기반 맞춤 제품 추천',
                '명동 올리브영 메가스토어',
                '명동 인근 약국',
            ],
            'note': '💳 쇼핑비 개별 결제',
        },
        {
            'time': '13:30',
            'title': '피부진단 & 에스테틱',
            'details': [
                '전문 에스테티션 1:1 피부 진단',
                '기초 페이셜 클렌징 / 수분 앰플 집중 관리 / 블랙헤드 & 모공 관리 중 택 1',
            ],
            'note': '💳 시술비 개별 결제',
        },
        {
            'time': '15:30',
            'title': '헤어 & 메이크업',
            'details': [
                '명동 제휴 헤어숍 스타일링',
                '한국식 글로우 메이크업 풀코스',
            ],
            'note': '✅ 포함',
        },
        {
            'time': '17:30',
            'title': '프로필 촬영',
            'details': [
                '명동 인근 실내 스튜디오',
                '전문 포토그래퍼 동행',
                '한국식 프로필 컷',
            ],
            'note': '✅ 포함',
        },
        {
            'time': '19:30',
            'title': '해산',
            'details': [],
            'note': '명동역 인근',
        },
        {
            'time': '별도 날짜',
            'title': '명동 TOP5 병원 시술 예약 대행',
            'details': [
                '원하는 시술 + 날짜 선택',
                '피부과 · 성형외과 연결',
                '시술비는 병원 현장 결제',
            ],
            'note': '🏥 예약 대행',
        },
    ]

    table = doc.add_table(rows=1 + len(schedule), cols=3)
    table.style = 'Table Grid'

    # 헤더
    for i, h in enumerate(['시간', '코스', '안내']):
        cell = table.rows[0].cells[i]
        cell.text = h
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].runs[0].font.size = Pt(10)
        set_cell_bg(cell, header_bg)

    for r_idx, item in enumerate(schedule):
        row = table.rows[r_idx + 1]

        # 시간 셀
        row.cells[0].text = item['time']
        row.cells[0].paragraphs[0].runs[0].font.size = Pt(10)

        # 코스 셀 — 제목 bold + 세부 항목
        course_cell = row.cells[1]
        course_cell.paragraphs[0].clear()
        title_run = course_cell.paragraphs[0].add_run(item['title'])
        title_run.bold = True
        title_run.font.size = Pt(10)
        for detail in item['details']:
            p = course_cell.add_paragraph('· ' + detail)
            p.runs[0].font.size = Pt(9)
            p.paragraph_format.space_before = Pt(1)

        # 안내 셀
        row.cells[2].text = item['note']
        row.cells[2].paragraphs[0].runs[0].font.size = Pt(10)

    return table

# ─────────────────────────────────────────────
# 고객공유용
# ─────────────────────────────────────────────
doc1 = Document()
doc1.core_properties.title = 'Seoul Glow Day — 당일 변신 패키지'

style = doc1.styles['Normal']
style.font.name = 'Malgun Gothic'
style.font.size = Pt(10)

add_heading(doc1, '✨ Seoul Glow Day — 당일 변신 패키지', 1, (31, 73, 125))
p = doc1.add_paragraph('하루 안에 피부·외모 변화를 경험하는 K-뷰티 데이트립')
p.runs[0].italic = True
p2 = doc1.add_paragraph('명동 전 구간 도보 이동 | 소그룹 운영')
p2.runs[0].italic = True

# 하루 일정 표
doc1.add_paragraph()
add_heading(doc1, '🗓️ 하루 일정', 2)
add_schedule_table(doc1, 'D9E2F3')

# 여행 편의 키트
doc1.add_paragraph()
add_heading(doc1, '🎁 여행 편의 키트', 2)
kit_items = [
    '한국 eSIM (여행 기간 데이터)',
    'T-money 교통카드',
    '여행자 보험',
    '코스별 예약 확인서',
    '카카오톡 긴급 연락 (한국어 · 영어)',
]
for item in kit_items:
    doc1.add_paragraph(item, style='List Bullet')

# 가격 안내
doc1.add_paragraph()
add_heading(doc1, '💰 가격 안내', 2)
p = doc1.add_paragraph()
p.add_run('기본 패키지  ₩399,000').bold = True

doc1.add_paragraph()
add_table(doc1,
    ['큐레이터 동행 구간', '추가 요금', '총 금액'],
    [
        ['없음', '—', '₩399,000'],
        ['1구간', '+₩50,000', '₩449,000'],
        ['2구간', '+₩80,000', '₩479,000'],
        ['3구간', '+₩110,000', '₩509,000'],
        ['4구간', '+₩130,000', '₩529,000'],
        ['5구간 (전 구간)', '+₩150,000', '₩549,000'],
    ],
    'D9E2F3'
)
doc1.add_paragraph('※ 큐레이터는 원하는 구간만 연속으로 선택 가능합니다').italic = True
doc1.add_paragraph('※ 지원 언어: 영어 · 중국어 · 일본어').italic = True

doc1.save('/home/user/seolle/packages/고객공유용_당일변신패키지.docx')
print('고객공유용 완료')

# ─────────────────────────────────────────────
# 내부공유용
# ─────────────────────────────────────────────
doc2 = Document()
doc2.core_properties.title = '[내부용] Seoul Glow Day 원가 & 수익 구조'

style2 = doc2.styles['Normal']
style2.font.name = 'Malgun Gothic'
style2.font.size = Pt(10)

add_heading(doc2, '[내부용] Seoul Glow Day — 원가 & 수익 구조', 1, (192, 0, 0))
p = doc2.add_paragraph('🔒 내부 공유 전용 — 외부 배포 금지')
p.runs[0].bold = True

doc2.add_paragraph()
add_heading(doc2, '[1] 패키지 코스 구성', 2)
add_table(doc2,
    ['시간', '코스', '포함 여부', '결제 주체'],
    [
        ['10:00', '퍼스널컬러 분석', '✅ 포함', '우리'],
        ['11:00', '올리브영 & 약국 쇼핑', '❌ 미포함', '고객 현장'],
        ['13:30', '피부진단 & 에스테틱', '❌ 미포함', '고객 현장'],
        ['15:30', '헤어 & 메이크업', '✅ 포함', '우리'],
        ['17:30', '프로필 촬영', '✅ 포함', '우리'],
        ['19:30', '해산', '—', '—'],
        ['별도 날짜', '명동 TOP5 병원 예약 대행', '예약만 우리', '고객 현장'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[2] 구성 상품별 원가', 2)
add_heading(doc2, '실체 상품 (핵심 3개)', 3)
add_table(doc2,
    ['항목', 'B2B 원가', '비고'],
    [
        ['퍼스널컬러 분석', '₩60,000', '예상 원가 (협약 전)'],
        ['헤어 & 메이크업', '₩70,000', '협약 전 예상가'],
        ['프로필 촬영', '₩50,000', '협약 완료'],
        ['소계', '₩180,000', ''],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '번들 상품 (원가율 낮추기용)', 3)
add_table(doc2,
    ['항목', '원가', '고객 인식 가치', '원가율'],
    [
        ['여행자 보험', '₩2,500', '₩12,000', '21%'],
        ['eSIM', '₩10,000', '₩30,000', '33%'],
        ['T-money 교통카드', '₩2,500', '₩10,000', '25%'],
        ['카카오톡 긴급 연락', '₩0', '₩5,000', '0%'],
        ['면세 쿠폰', '₩0', '₩5,000', '0%'],
        ['예약 확인서 발송', '₩0', '₩3,000', '0%'],
        ['소계', '₩15,000', '₩65,000', '23%'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[3] 원가 합계', 2)
add_table(doc2,
    ['구분', '금액'],
    [
        ['실체 상품', '₩180,000'],
        ['번들 상품', '₩15,000'],
        ['운영비', '₩15,000'],
        ['총 원가', '₩210,000'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[4] 패키지 판매가 & 원가율', 2)
add_table(doc2,
    ['판매가', '원가', '원가율', '마진'],
    [
        ['₩350,000', '₩210,000', '60.0%', '₩140,000'],
        ['₩380,000', '₩210,000', '55.3%', '₩170,000'],
        ['₩399,000 ✅ 적용', '₩210,000', '52.6%', '₩189,000'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[5] 큐레이터 수익 구조 (중개 수수료 20%)', 2)
add_table(doc2,
    ['구간', '고객 지불', '큐레이터 지급 (80%)', '우리 수익 (20%)'],
    [
        ['1구간', '₩50,000', '₩40,000', '₩10,000'],
        ['2구간', '₩80,000', '₩64,000', '₩16,000'],
        ['3구간', '₩110,000', '₩88,000', '₩22,000'],
        ['4구간', '₩130,000', '₩104,000', '₩26,000'],
        ['5구간', '₩150,000', '₩120,000', '₩30,000'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[6] 큐레이터 포함 시 총 수익 & 전체 원가율 (기본 ₩399,000 기준)', 2)
add_table(doc2,
    ['큐레이터 구간', '총 판매가', '총 원가', '전체 원가율', '총 마진'],
    [
        ['없음', '₩399,000', '₩210,000', '52.6%', '₩189,000'],
        ['1구간', '₩449,000', '₩250,000', '55.7%', '₩199,000'],
        ['2구간', '₩479,000', '₩274,000', '57.2%', '₩205,000'],
        ['3구간', '₩509,000', '₩298,000', '58.5%', '₩211,000'],
        ['4구간', '₩529,000', '₩314,000', '59.4%', '₩215,000'],
        ['5구간', '₩549,000', '₩330,000', '60.1% ⚠️', '₩219,000'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[7] 현장결제 항목 수익 검토', 2)
add_table(doc2,
    ['항목', '고객 결제', '우리 수익'],
    [
        ['올리브영 · 약국 쇼핑', '고객 직접', '면세쿠폰 리베이트 가능성 검토'],
        ['피부진단 & 에스테틱', '고객 직접', '예약 대행 수수료 협약 검토'],
        ['명동 TOP5 병원 시술', '고객 직접', '예약 대행 수수료 협약 검토'],
    ],
    'FCE4D6'
)

doc2.add_paragraph()
add_heading(doc2, '[8] 협약 미완료 항목', 2)
add_table(doc2,
    ['항목', '상태', '비고'],
    [
        ['퍼스널컬러 전문가', '🔴 미완료', '원가 ₩60,000 예상'],
        ['헤어 & 메이크업', '🔴 미완료', '원가 ₩70,000 예상'],
        ['프로필 촬영', '🟢 완료', '₩50,000 확정'],
        ['에스테틱 업체', '🔴 미완료', '예약 대행 수수료 협의 필요'],
        ['명동 TOP5 병원', '🔴 미완료', '협약 + 외국인 특별가 확인'],
        ['eSIM 공급처', '🔴 미완료', '도매 or 대리점 제휴'],
        ['여행자 보험사', '🔴 미완료', '대행 수수료 구조 확인'],
    ],
    'FCE4D6'
)

doc2.save('/home/user/seolle/packages/내부공유용_당일변신패키지.docx')
print('내부공유용 완료')
