# 품질 점검용 합성 문서 만들기(고객 문서 없음). 문서마다 기능 몇 개씩 — 차이가 나면 원인을 바로 짚을 수 있게.
# 사용: python audit/office/gen_corpus.py <출력폴더>  → <출력>/raw/{docx,xlsx,pptx}/*
# 그 뒤 ref_office.ps1이 MS Office로 열어 다시 저장(실제 사용자 파일과 같은 모양)하고 PDF(정답)를 만든다.
import io, os, sys, datetime, random
from PIL import Image, ImageDraw
from lxml import etree

OUT = os.path.join(sys.argv[1] if len(sys.argv) > 1 else 'audit/office/corpus', 'raw')
for k in ('docx', 'xlsx', 'pptx'):
    os.makedirs(os.path.join(OUT, k), exist_ok=True)


def png(w, h, kind='photo'):
    if kind == 'logo':
        im = Image.new('RGB', (w, h), (255, 255, 255)); d = ImageDraw.Draw(im)
        d.rounded_rectangle([2, 2, w - 3, h - 3], radius=12, fill=(20, 70, 160))
        d.rectangle([w * 0.15, h * 0.3, w * 0.85, h * 0.7], fill=(255, 255, 255))
    else:
        im = Image.new('RGB', (w, h)); d = ImageDraw.Draw(im)
        for y in range(h):  # 하늘색 → 초록 그라데이션(사진 대용)
            t = y / h
            d.line([(0, y), (w, y)], fill=(int(90 + 120 * t), int(150 + 60 * t), int(220 - 150 * t)))
        d.ellipse([w * 0.6, h * 0.1, w * 0.85, h * 0.4], fill=(255, 210, 60))
        d.polygon([(0, h), (w * 0.35, h * 0.45), (w * 0.7, h)], fill=(60, 110, 60))
        d.rectangle([4, 4, w - 5, h - 5], outline=(200, 0, 0), width=6)  # 빨간 테두리: 잘림 확인용
        d.line([(0, 0), (w, h)], fill=(0, 0, 0), width=4)  # 대각선: 늘어남·자르기 확인용
    b = io.BytesIO(); im.save(b, 'PNG'); b.seek(0); return b


KO = '향기랩 주식회사는 2026년 하반기 사업 계획에 따라 설비 점검과 품질 관리 체계를 개선한다. '
EN = 'The quick brown fox jumps over the lazy dog 0123456789. '

# ───────────────────────── DOCX ─────────────────────────
import docx
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_BREAK, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def set_font(run, name, size=None, color=None, bold=None):
    run.font.name = name
    rpr = run._element.get_or_add_rPr()
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = rpr.makeelement(qn('w:rFonts'), {}); rpr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:eastAsia', 'w:cs'):
        rf.set(qn(a), name)
    if size: run.font.size = Pt(size)
    if color: run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None: run.bold = bold


def new_doc():
    d = docx.Document()
    s = d.sections[0]; s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin = s.right_margin = Cm(2.5); s.top_margin = s.bottom_margin = Cm(2.5)
    st = d.styles['Normal']; st.font.name = '맑은 고딕'; st.font.size = Pt(10.5)
    st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    return d


def d01():
    d = new_doc()
    d.add_heading('1. 글꼴·문단 서식 시험', level=1)
    for font in ('맑은 고딕', '바탕', '굴림', '돋움', 'Times New Roman', 'Arial'):
        p = d.add_paragraph(); r = p.add_run(f'[{font}] '); set_font(r, font, 11, bold=True)
        r = p.add_run('가나다라마바사 한글 본문 ABC abc 123 — ' + EN[:30]); set_font(r, font, 11)
    p = d.add_paragraph()
    for txt in ('굵게 ', '기울임 ', '밑줄 ', '취소선 ', '빨강 ', '형광펜'):
        r = p.add_run(txt)
        if txt.startswith('굵'): r.bold = True
        if txt.startswith('기울'): r.italic = True
        if txt.startswith('밑'): r.underline = True
        if txt.startswith('취소'): r.font.strike = True
        if txt.startswith('빨'): r.font.color.rgb = RGBColor.from_string('C00000')
        if txt == '형광펜': r.font.highlight_color = 7
    p = d.add_paragraph(); p.add_run('위첨자 x'); r2 = p.add_run('2'); r2.font.superscript = True
    p.add_run(' 아래첨자 H'); r3 = p.add_run('2'); r3.font.subscript = True; p.add_run('O, 글자 크기 ')
    for sz in (8, 12, 16, 20, 28):
        rr = p.add_run(f'{sz}pt '); rr.font.size = Pt(sz)
    for label, rule, val in [('줄 간격 1.0', WD_LINE_SPACING.SINGLE, None), ('줄 간격 1.5', WD_LINE_SPACING.ONE_POINT_FIVE, None),
                             ('줄 간격 2.0', WD_LINE_SPACING.DOUBLE, None), ('줄 간격 고정 14pt', WD_LINE_SPACING.EXACTLY, Pt(14))]:
        d.add_heading(label, level=2)
        p = d.add_paragraph(KO * 3); pf = p.paragraph_format; pf.line_spacing_rule = rule
        if val: pf.line_spacing = val
        pf.space_after = Pt(12)
    for al, name in [(WD_ALIGN_PARAGRAPH.LEFT, '왼쪽'), (WD_ALIGN_PARAGRAPH.CENTER, '가운데'), (WD_ALIGN_PARAGRAPH.RIGHT, '오른쪽'), (WD_ALIGN_PARAGRAPH.JUSTIFY, '양쪽')]:
        p = d.add_paragraph(f'[{name} 맞춤] ' + KO * 2); p.alignment = al
    p = d.add_paragraph('[첫 줄 들여쓰기 1cm, 왼쪽 2cm] ' + KO * 2); p.paragraph_format.first_line_indent = Cm(1); p.paragraph_format.left_indent = Cm(2)
    p = d.add_paragraph('[내어쓰기] ' + KO * 2); p.paragraph_format.first_line_indent = Cm(-1); p.paragraph_format.left_indent = Cm(1)
    d.add_paragraph('부제목 스타일 예시', style='Subtitle'); d.add_paragraph('인용 스타일 예시 — ' + KO, style='Quote')
    d.save(f'{OUT}/docx/d01-styles-fonts.docx')


NUMBERING_ABS = '<w:abstractNum %s w:abstractNumId="%d"><w:multiLevelType w:val="multilevel"/>%s</w:abstractNum>'


def lvl(i, fmt, text, ind):
    return (f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{ind}" w:hanging="425"/></w:pPr></w:lvl>')


def d02():
    d = new_doc()
    numbering = d.part.numbering_part.element
    levels = ''.join([lvl(0, 'decimal', '%1.', 425), lvl(1, 'decimal', '%1.%2.', 851), lvl(2, 'decimal', '%1.%2.%3.', 1276),
                      lvl(3, 'ganada', '%4)', 1701), lvl(4, 'chosung', '(%5)', 2126)])
    levels2 = ''.join([lvl(0, 'bullet', '●', 425), lvl(1, 'bullet', '○', 851), lvl(2, 'bullet', '■', 1276), lvl(3, 'bullet', '-', 1701)])
    numbering.insert(0, parse_xml(NUMBERING_ABS % (nsdecls('w'), 91, levels2)))
    numbering.insert(0, parse_xml(NUMBERING_ABS % (nsdecls('w'), 90, levels)))
    for n in (90, 91):
        numbering.append(parse_xml(f'<w:num {nsdecls("w")} w:numId="{n}"><w:abstractNumId w:val="{n}"/></w:num>'))

    def item(text, num, il):
        p = d.add_paragraph(text); ppr = p._p.get_or_add_pPr()
        ppr.append(parse_xml(f'<w:numPr {nsdecls("w")}><w:ilvl w:val="{il}"/><w:numId w:val="{num}"/></w:numPr>'))
    d.add_heading('다단계 번호 매기기', 1)
    seq = [(0, '사업 개요'), (1, '추진 배경'), (2, '설비 노후화에 따른 고장 증가'), (2, '법정 점검 주기 강화'), (3, '전기 설비'), (3, '기계 설비'),
           (4, '공조기'), (4, '펌프'), (1, '추진 목표'), (0, '세부 계획'), (1, '1단계: 현황 조사 — ' + KO), (1, '2단계: 개선 공사'), (0, '예산')]
    for il, t in seq: item(t, 90, il)
    d.add_heading('글머리 기호', 1)
    for il, t in [(0, '첫째 수준 글머리'), (1, '둘째 수준'), (2, '셋째 수준 ' + KO), (3, '넷째 수준'), (0, '다시 첫째')]: item(t, 91, il)
    d.add_heading('기본 스타일 목록(List Number / List Bullet)', 1)
    for t in ('가', '나', '다'): d.add_paragraph('List Number ' + t, style='List Number')
    for t in ('가', '나'): d.add_paragraph('List Bullet ' + t, style='List Bullet')
    d.save(f'{OUT}/docx/d02-numbering.docx')


def shade(cell, hexcolor):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{hexcolor}"/>'))


def borders(cell, **kw):
    tcPr = cell._tc.get_or_add_tcPr(); b = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for side, (val, sz, color) in kw.items():
        b.append(parse_xml(f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" w:color="{color}"/>'))
    tcPr.append(b)


def d03():
    d = new_doc()
    d.add_heading('표: 병합·테두리·음영', 1)
    t = d.add_table(rows=7, cols=5); t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = [Cm(2), Cm(4.5), Cm(3), Cm(3), Cm(3.5)]
    for row in t.rows:
        for i, c in enumerate(row.cells): c.width = widths[i]
    for i, h in enumerate(['구분', '항목', '1분기', '2분기', '비고']):
        c = t.cell(0, i); c.text = h; shade(c, '1F4E79'); r = c.paragraphs[0].runs[0]; r.font.color.rgb = RGBColor(255, 255, 255); r.bold = True
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    data = [['설비', '공조기 점검', '1,200', '1,350', ''], ['', '펌프 교체', '3,400', '0', '긴급'], ['', '배관 보수', '800', '920', ''],
            ['전기', '수변전 설비', '2,100', '2,300', '법정 점검'], ['', '조명 LED 교체 공사(본관 1~5층 전체)', '5,600', '4,800', ''], ['합계', '', '13,100', '9,370', '']]
    for r, row in enumerate(data, 1):
        for c, v in enumerate(row):
            if v: t.cell(r, c).text = v
            if c >= 2: t.cell(r, c).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    t.cell(1, 0).merge(t.cell(3, 0)); t.cell(4, 0).merge(t.cell(5, 0)); t.cell(6, 0).merge(t.cell(6, 1))
    t.cell(1, 0).vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    for c in range(5): shade(t.cell(6, c), 'FFF2CC')
    shade(t.cell(2, 4), 'F8CBAD')
    d.add_paragraph()
    d.add_heading('테두리 종류', 2)
    t2 = d.add_table(rows=2, cols=4)
    for i, (val, sz, col) in enumerate([('single', 4, '000000'), ('double', 6, 'C00000'), ('dashed', 8, '0070C0'), ('thick', 24, '00B050')]):
        for r in range(2):
            c = t2.cell(r, i); c.text = f'{val} {sz}' if r == 0 else '내용'
            borders(c, top=(val, sz, col), left=(val, sz, col), bottom=(val, sz, col), right=(val, sz, col))
    d.add_paragraph()
    d.add_heading('긴 표(쪽 넘김·제목 행 반복)', 2)
    t3 = d.add_table(rows=1, cols=3); t3.style = 'Light Grid Accent 1'
    for i, h in enumerate(['번호', '설비명', '점검 결과']): t3.cell(0, i).text = h
    t3.rows[0]._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for n in range(1, 41):
        row = t3.add_row().cells; row[0].text = str(n); row[1].text = f'설비 {n:03d}호기'; row[2].text = ['양호', '불량', '보류'][n % 3]
    d.save(f'{OUT}/docx/d03-tables.docx')


ANCHOR = ('<wp:anchor %s distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" '
          'layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>'
          '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="%d" cy="%d"/>'
          '<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>%s<wp:cNvGraphicFramePr/>%s</wp:anchor>')
WP_NS = ('xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
         'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def float_picture(inline):
    ext = inline.find(qn('wp:extent')); docPr = inline.find(qn('wp:docPr')); graphic = inline.find(qn('a:graphic'))
    anchor = parse_xml(ANCHOR % (WP_NS, int(ext.get('cx')), int(ext.get('cy')), etree.tostring(docPr).decode(), etree.tostring(graphic).decode()))
    inline.getparent().replace(inline, anchor)


def d04():
    d = new_doc()
    d.add_heading('그림: 줄 안·떠 있는 그림', 1)
    d.add_paragraph('아래는 줄 안(inline) 그림 12cm 폭입니다.')
    d.add_picture(png(800, 450), width=Cm(12))
    d.add_paragraph('그림 1. 줄 안 그림 캡션').alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = d.add_paragraph(); r = p.add_run(); r.add_picture(png(400, 300), width=Cm(6))
    float_picture(r._element.find('.//' + qn('wp:inline')))
    p.add_run('오른쪽에 떠 있는 그림(정사각형 배치) 주변으로 글이 흘러야 합니다. ' + KO * 6)
    d.add_paragraph(KO * 4)
    p = d.add_paragraph('글자 사이 작은 아이콘 '); r = p.add_run(); r.add_picture(png(60, 60, 'logo'), width=Cm(0.6)); p.add_run(' 이 글줄 높이에 맞게 있어야 함.')
    d.save(f'{OUT}/docx/d04-images.docx')


def add_field(par, instr):
    r = par.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>'))
    r = par.add_run(); r._r.append(parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> {instr} </w:instrText>'))
    r = par.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>'))
    par.add_run('1')
    r = par.add_run(); r._r.append(parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>'))


def d05():
    d = new_doc(); s = d.sections[0]; s.different_first_page_header_footer = True
    hp = s.header.paragraphs[0]; hp.text = '향기랩 설비관리팀\t\t대외비'
    s.first_page_header.paragraphs[0].text = '표지 머리글(첫 쪽만)'
    fp = s.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER; fp.add_run('- '); add_field(fp, 'PAGE'); fp.add_run(' / '); add_field(fp, 'NUMPAGES'); fp.add_run(' -')
    t = d.add_paragraph(); t.alignment = WD_ALIGN_PARAGRAPH.CENTER; r = t.add_run('2026년 설비 점검 결과 보고서'); r.font.size = Pt(26); r.bold = True
    d.add_paragraph('표지 — 이 쪽에는 일반 머리글이 없고 첫 쪽 머리글만 나와야 함').alignment = WD_ALIGN_PARAGRAPH.CENTER
    d.add_page_break()
    for i in range(1, 4):
        d.add_heading(f'{i}장 본문', 1)
        for _ in range(3): d.add_paragraph(KO * 4)
        if i < 3: d.add_page_break()
    s2 = d.add_section(WD_SECTION.NEW_PAGE); s2.orientation = WD_ORIENT.LANDSCAPE; s2.page_width, s2.page_height = Cm(29.7), Cm(21)
    d.add_heading('가로 방향 구역', 1); d.add_paragraph('이 쪽은 가로(A4 landscape)여야 합니다. ' + KO * 3)
    d.save(f'{OUT}/docx/d05-header-footer-sections.docx')


FOOTNOTES = ('<w:footnotes %s><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
             '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>%s</w:footnotes>')


def d06():
    from docx.opc.part import Part
    from docx.opc.packuri import PackURI
    d = new_doc()
    d.add_heading('2단 구성과 각주', 1)
    d.add_paragraph('이 문단까지는 1단입니다. ' + KO)
    d.add_section(WD_SECTION.CONTINUOUS)  # 앞 구역(1단) 끝
    notes = []
    for i in range(1, 4):
        p = d.add_paragraph(f'단락 {i}: ' + KO * 2)
        r = p.add_run(); r._r.append(parse_xml(f'<w:rPr {nsdecls("w")}><w:vertAlign w:val="superscript"/></w:rPr>'))
        r._r.append(parse_xml(f'<w:footnoteReference {nsdecls("w")} w:id="{i}"/>'))
        p.add_run(' ' + KO * 2)
        notes.append(f'<w:footnote w:id="{i}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
                     f'<w:r><w:t xml:space="preserve"> 각주 {i}: 관련 법령 제{i}조 참조. 설비 점검 기준 고시.</w:t></w:r></w:p></w:footnote>')
    for _ in range(4): d.add_paragraph(KO * 3)
    s2 = d.add_section(WD_SECTION.CONTINUOUS)  # 이 sectPr이 2단 구역을 끝낸다
    prev = d.sections[-2]._sectPr
    prev.append(parse_xml(f'<w:cols {nsdecls("w")} w:num="2" w:space="425" w:sep="1"/>'))
    d.add_paragraph('다시 1단으로 돌아온 마지막 문단.')
    xml = (FOOTNOTES % (nsdecls('w'), ''.join(notes))).encode('utf-8')
    part = Part(PackURI('/word/footnotes.xml'), 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml', xml, d.part.package)
    d.part.relate_to(part, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes')
    d.save(f'{OUT}/docx/d06-columns-footnotes.docx')


TEXTBOX = ('<w:r %s><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251660288" behindDoc="0" '
           'locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>3200400</wp:posOffset></wp:positionH>'
           '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2286000" cy="1143000"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
           '<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="50" name="TextBox 1"/><wp:cNvGraphicFramePr/><a:graphic>'
           '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
           '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2286000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
           '<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="ED7D31"/></a:solidFill></a:ln></wps:spPr>'
           '<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>글상자: 주의 사항</w:t></w:r></w:p>'
           '<w:p><w:r><w:t>점검 전 전원 차단을 확인할 것.</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
           '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')
TB_NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
         'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')


def d07():
    d = new_doc()
    d.add_heading('탭 정지·글상자·쪽 나누기', 1)
    p = d.add_paragraph(); ts = p.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(8), WD_TAB_ALIGNMENT.CENTER); ts.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    p.add_run('왼쪽\t가운데\t오른쪽(점선 채움)')
    for name, pg in [('제1장 사업 개요', 1), ('제2장 추진 계획', 3), ('제3장 예산 및 일정', 12), ('부록', 27)]:
        p = d.add_paragraph(); p.paragraph_format.tab_stops.add_tab_stop(Cm(16), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS); p.add_run(f'{name}\t{pg}')
    for v in ('1,234.5', '12.75'):
        p = d.add_paragraph(); p.paragraph_format.tab_stops.add_tab_stop(Cm(10), WD_TAB_ALIGNMENT.DECIMAL); p.add_run(f'소수점 맞춤\t{v}')
    p = d.add_paragraph('글상자가 오른쪽에 떠 있어야 합니다. ')
    p._p.append(parse_xml(TEXTBOX % TB_NS)); p.add_run(KO * 5)
    p = d.add_paragraph('이 문장 뒤에서 쪽을 나눕니다.'); p.add_run().add_break(WD_BREAK.PAGE)
    d.add_paragraph('둘째 쪽 첫 줄 — 위 쪽 나누기 뒤에 와야 합니다.')
    p = d.add_paragraph('문단 앞 쪽 나누기 속성(pageBreakBefore)이 있는 문단 — 셋째 쪽 맨 위'); p.paragraph_format.page_break_before = True
    d.save(f'{OUT}/docx/d07-tabs-textbox-breaks.docx')


def d08():
    d = new_doc()
    t = d.add_paragraph(); r = t.add_run('공 문'); set_font(r, '바탕', 22, bold=True); t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info = d.add_table(rows=4, cols=2); info.style = 'Table Grid'
    for i, (k, v) in enumerate([('수신', '각 부서장'), ('발신', '향기랩 설비관리팀장'), ('제목', '2026년 하반기 정기 점검 일정 안내'), ('시행일', '2026. 11. 20.')]):
        info.cell(i, 0).text = k; info.cell(i, 1).text = v; shade(info.cell(i, 0), 'D9E2F3'); info.cell(i, 0).width = Cm(3); info.cell(i, 1).width = Cm(13)
    d.add_paragraph()
    for txt in ('1. 귀 부서의 무궁한 발전을 기원합니다.',
                '2. 「시설물의 안전 및 유지관리에 관한 특별법」 제11조에 따라 아래와 같이 정기 점검을 시행하오니 협조하여 주시기 바랍니다.'):
        p = d.add_paragraph(txt); p.paragraph_format.line_spacing = 1.6
    for txt in ('  가. 점검 기간: 2026. 12. 1.(월) ~ 12. 12.(금)', '  나. 점검 대상: 본관·별관 전기, 기계, 소방 설비', '  다. 협조 사항: 점검 시간 중 정전(약 30분) 예정'):
        d.add_paragraph(txt)
    d.add_paragraph('붙임  1. 점검 일정표 1부.  끝.')
    p = d.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; r = p.add_run('향 기 랩 주 식 회 사 장'); set_font(r, '바탕', 18, bold=True)
    p = d.add_paragraph(); r = p.add_run(); r.add_picture(png(200, 80, 'logo'), width=Cm(3)); p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    d.save(f'{OUT}/docx/d08-korean-official-letter.docx')


# ───────────────────────── XLSX ─────────────────────────
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, IconSetRule, CellIsRule, FormulaRule
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XLImage
thin = Side(style='thin', color='000000'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)


def wb_new(title):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = title; return wb, ws


def x01():
    wb, ws = wb_new('숫자서식')
    rows = [('정수 #,##0', 1234567, '#,##0'), ('소수 0.00', 3.14159, '0.00'), ('원화', 1500000, '"₩"#,##0'),
            ('회계', -25000, '_-"₩"* #,##0_-;-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-'), ('음수 빨강', -4321, '#,##0;[Red]-#,##0'),
            ('음수 괄호', -4321, '#,##0;(#,##0)'), ('백분율', 0.256, '0.0%'), ('지수', 123456789, '0.00E+00'), ('분수', 0.75, '# ?/?'),
            ('날짜', datetime.date(2026, 11, 20), 'yyyy-mm-dd'), ('날짜 한글', datetime.date(2026, 11, 20), 'yyyy"년" m"월" d"일"'),
            ('요일', datetime.date(2026, 11, 20), 'yyyy-mm-dd (aaa)'), ('시간', datetime.time(14, 5, 9), 'hh:mm:ss'),
            ('오전/오후', datetime.time(14, 5), 'AM/PM h:mm'), ('날짜+시간', datetime.datetime(2026, 11, 20, 9, 30), 'yyyy-mm-dd hh:mm'),
            ('텍스트 @', '00123', '@'), ('앞 0 채움', 42, '00000'), ('천 단위 생략', 12345678, '#,##0,"천원"'),
            ('조건 서식코드', 1500, '[>=1000]"많음 "#,##0;"적음 "#,##0'), ('USD', 1234.5, '[$$-409]#,##0.00'), ('일반 0.1+0.2', 0.1 + 0.2, 'General')]
    ws['A1'] = '서식'; ws['B1'] = '값'; ws['C1'] = '서식 코드'
    for i, (n, v, f) in enumerate(rows, 2):
        ws.cell(i, 1, n); c = ws.cell(i, 2, v); c.number_format = f; ws.cell(i, 3, f)
    ws.column_dimensions['A'].width = 16; ws.column_dimensions['B'].width = 22; ws.column_dimensions['C'].width = 44
    ws.column_dimensions['E'].width = 6; ws['D2'] = '좁은 칸 →'; ws['E2'] = 123456789012; ws['E2'].number_format = '#,##0'  # #### 표시 확인
    wb.save(f'{OUT}/xlsx/x01-number-formats.xlsx')


def x02():
    wb, ws = wb_new('병합테두리')
    ws.merge_cells('A1:F1'); ws['A1'] = '2026년 분기별 설비 예산(단위: 천원)'; ws['A1'].font = Font(name='맑은 고딕', size=16, bold=True, color='1F4E79')
    ws['A1'].alignment = Alignment(horizontal='center', vertical='center'); ws.row_dimensions[1].height = 32
    ws.merge_cells('A2:A3'); ws.merge_cells('B2:B3'); ws.merge_cells('C2:F2')
    for c, v in {'A2': '구분', 'B2': '항목', 'C2': '분기', 'C3': '1분기', 'D3': '2분기', 'E3': '3분기', 'F3': '4분기'}.items(): ws[c] = v
    hdr = PatternFill('solid', fgColor='D9E2F3')
    for row in ws['A2:F3']:
        for c in row: c.fill = hdr; c.border = BOX; c.font = Font(bold=True); c.alignment = Alignment(horizontal='center', vertical='center')
    data = [('설비', '공조기', 1200, 1350, 900, 1100), (None, '펌프', 3400, 0, 500, 200), ('전기', '수변전', 2100, 2300, 2000, 2500), (None, '조명 LED 교체 공사 본관 전층', 5600, 4800, 3000, 1000)]
    for r, row in enumerate(data, 4):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v); cell.border = BOX
            if c >= 3: cell.number_format = '#,##0'
    ws.merge_cells('A4:A5'); ws.merge_cells('A6:A7'); ws['A4'].alignment = ws['A6'].alignment = Alignment(vertical='center', horizontal='center')
    ws['B7'].alignment = Alignment(wrap_text=True, vertical='top')
    for i, st in enumerate(['thin', 'medium', 'thick', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot']):
        c = ws.cell(10, 1 + i, st); s = Side(style=st, color='C00000'); c.border = Border(left=s, right=s, top=s, bottom=s)
    fills = [PatternFill('solid', fgColor='FFC000'), PatternFill('darkGrid', fgColor='4472C4', bgColor='FFFFFF'), PatternFill('lightTrellis', fgColor='70AD47'), PatternFill('gray125')]
    for i, f in enumerate(fills): c = ws.cell(12, 1 + i, f.fill_type); c.fill = f
    ws['A14'] = '세로쓰기'; ws['A14'].alignment = Alignment(text_rotation=255); ws['B14'] = '45도 회전'; ws['B14'].alignment = Alignment(text_rotation=45)
    ws['C14'] = '들여쓰기2'; ws['C14'].alignment = Alignment(indent=2)
    ws['D14'] = '셀에 맞춤 축소 글자가 아주 깁니다'; ws['D14'].alignment = Alignment(shrink_to_fit=True)
    ws['E14'] = '자동 줄 바꿈이 켜진 긴 글자 칸입니다 여러 줄로 보여야 함'; ws['E14'].alignment = Alignment(wrap_text=True)
    ws['F14'] = '넘침: 옆 칸이 비어 있으면 글자가 오른쪽 칸으로 넘쳐 보여야 합니다'
    ws.row_dimensions[14].height = 60
    for col, w in zip('ABCDEFGH', [8, 22, 10, 10, 14, 10, 10, 14]): ws.column_dimensions[col].width = w
    for i, (fn, sz) in enumerate([('맑은 고딕', 11), ('굴림', 11), ('바탕', 11), ('돋움', 9), ('Arial', 14), ('맑은 고딕', 20)]):
        c = ws.cell(17, 1 + i, f'{fn}{sz}가A1'); c.font = Font(name=fn, size=sz, italic=(i == 4), underline='single' if i == 3 else None, color='7030A0' if i == 5 else None)
    wb.save(f'{OUT}/xlsx/x02-merge-border-fill.xlsx')


def x03():
    wb, ws = wb_new('조건부서식')
    ws.append(['지점', '매출', '달성률', '증감', '평가'])
    random.seed(7)
    for i in range(1, 16): ws.append([f'지점{i:02d}', random.randint(100, 1000) * 10, round(random.uniform(0.4, 1.3), 2), random.randint(-50, 50), random.randint(1, 100)])
    ws.conditional_formatting.add('B2:B16', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
    ws.conditional_formatting.add('C2:C16', DataBarRule(start_type='min', end_type='max', color='5B9BD5', showValue=True))
    ws.conditional_formatting.add('E2:E16', IconSetRule('3TrafficLights1', 'percent', [0, 33, 67]))
    ws.conditional_formatting.add('D2:D16', CellIsRule(operator='lessThan', formula=['0'], font=Font(color='9C0006'), fill=PatternFill('solid', bgColor='FFC7CE')))
    ws.conditional_formatting.add('A2:A16', FormulaRule(formula=['$C2>=1'], font=Font(bold=True, color='006100'), fill=PatternFill('solid', bgColor='C6EFCE')))
    for r in range(2, 17): ws.cell(r, 3).number_format = '0%'; ws.cell(r, 2).number_format = '#,##0'
    ws['A18'] = '아이콘 5단계(화살표)'
    for i in range(5): ws.cell(19, 1 + i, (i + 1) * 20)
    ws.conditional_formatting.add('A19:E19', IconSetRule('5Arrows', 'percent', [0, 20, 40, 60, 80]))
    for col in 'ABCDE': ws.column_dimensions[col].width = 14
    wb.save(f'{OUT}/xlsx/x03-conditional-formatting.xlsx')


def x04():
    wb, ws = wb_new('틀고정')
    ws.append(['번호', '설비명', '위치', '숨긴열', '상태', '점검일', '메모'])
    for i in range(1, 61):
        ws.append([i, f'설비 {i:03d}', ['본관', '별관', '지하'][i % 3], 'X', ['양호', '불량'][i % 2], datetime.date(2026, 1, 1) + datetime.timedelta(days=i * 3), '자동 줄 바꿈 메모 ' * (1 + i % 3)])
    ws.freeze_panes = 'C2'; ws.column_dimensions['D'].hidden = True
    for r in (5, 6, 7): ws.row_dimensions[r].hidden = True
    for col, w in zip('ABCDEFG', [6, 30, 8, 8, 10, 12, 25]): ws.column_dimensions[col].width = w
    for r in range(2, 62): ws.cell(r, 7).alignment = Alignment(wrap_text=True, vertical='top'); ws.cell(r, 6).number_format = 'yyyy-mm-dd'
    ws.row_dimensions[2].height = 45; ws.row_dimensions[3].height = 8
    ws.sheet_view.zoomScale = 85
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = PatternFill('solid', fgColor='305496')
    ws2 = wb.create_sheet('격자선없음'); ws2.sheet_view.showGridLines = False; ws2['B2'] = '격자선이 보이면 안 되는 시트'; ws2['B4'] = '테두리만'; ws2['B4'].border = BOX
    ws2.sheet_properties.tabColor = 'FF0000'
    wb.save(f'{OUT}/xlsx/x04-freeze-hidden-sizes.xlsx')


def x05():
    wb, ws = wb_new('수식')
    ws.append(['품목', '단가', '수량', '금액', '부가세', '합계'])
    items = [('볼트', 120, 500), ('너트', 80, 500), ('배관 50A', 15000, 12), ('밸브', 48000, 3), ('필터', 22000, 8)]
    for i, (n, p, q) in enumerate(items, 2): ws.append([n, p, q, f'=B{i}*C{i}', f'=ROUND(D{i}*0.1,0)', f'=D{i}+E{i}'])
    n = len(items) + 1
    ws.append(['합계', None, f'=SUM(C2:C{n})', f'=SUM(D2:D{n})', f'=SUM(E2:E{n})', f'=SUM(F2:F{n})'])
    for r in range(2, n + 2):
        for c in (2, 4, 5, 6): ws.cell(r, c).number_format = '#,##0'
    for i, (k, f) in enumerate([('평균 단가', f'=AVERAGE(B2:B{n})'), ('최대', f'=MAX(F2:F{n})'), ('찾기(밸브)', f'=VLOOKUP("밸브",A2:F{n},6,FALSE)'),
                                ('조건', '=IF(I2>100000,"고액","일반")'), ('TEXT', '=TEXT(DATE(2026,11,20),"yyyy-mm-dd")'), ('오류', '=1/0'),
                                ('COUNTIF', f'=COUNTIF(C2:C{n},">=10")'), ('SUMPRODUCT', f'=SUMPRODUCT(B2:B{n},C2:C{n})'), ('다른 시트', "='참조'!B2*2")], 1):
        ws.cell(i, 8, k); ws.cell(i, 9, f).number_format = '#,##0.0' if i == 1 else 'General'
    ws2 = wb.create_sheet('참조'); ws2['A2'] = '기준값'; ws2['B2'] = 21
    for col, w in zip('ABCDEFGHI', [12, 10, 8, 12, 10, 12, 2, 14, 14]): ws.column_dimensions[col].width = w
    wb.save(f'{OUT}/xlsx/x05-formulas.xlsx')


def x06():
    wb, ws = wb_new('차트')
    ws.append(['월', '본관', '별관', '지하'])
    for i, m in enumerate(['1월', '2월', '3월', '4월', '5월', '6월'], 1): ws.append([m, 100 + i * 15, 80 + (i % 3) * 30, 50 + i * 5])
    bar = BarChart(); bar.title = '월별 전력 사용량(MWh)'; bar.y_axis.title = 'MWh'; bar.x_axis.title = '월'
    bar.add_data(Reference(ws, min_col=2, max_col=4, min_row=1, max_row=7), titles_from_data=True); bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7)); bar.width, bar.height = 16, 8
    ws.add_chart(bar, 'F1')
    line = LineChart(); line.title = '추세'; line.add_data(Reference(ws, min_col=2, max_col=3, min_row=1, max_row=7), titles_from_data=True)
    line.set_categories(Reference(ws, min_col=1, min_row=2, max_row=7)); line.width, line.height = 16, 8
    ws.add_chart(line, 'F18')
    ws['A10'] = '구분'; ws['B10'] = '비율'
    for i, (k, v) in enumerate([('전기', 45), ('기계', 30), ('소방', 15), ('기타', 10)], 11): ws.cell(i, 1, k); ws.cell(i, 2, v)
    pie = PieChart(); pie.title = '비용 구성'; pie.add_data(Reference(ws, min_col=2, min_row=10, max_row=14), titles_from_data=True)
    pie.set_categories(Reference(ws, min_col=1, min_row=11, max_row=14)); pie.width, pie.height = 10, 8
    pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
    ws.add_chart(pie, 'A17')
    wb.save(f'{OUT}/xlsx/x06-charts.xlsx')


def x07():
    wb, ws = wb_new('인쇄영역')
    for r in range(1, 41):
        for c in range(1, 13): ws.cell(r, c, f'R{r}C{c}' if (r + c) % 5 else r * c)
    ws.print_area = 'A1:H30'; ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.oddHeader.center.text = '인쇄 머리글 &D'; ws.oddFooter.center.text = '&P / &N 쪽'; ws.print_title_rows = '1:1'
    ws2 = wb.create_sheet('둘째시트'); ws2['A1'] = '둘째 시트 내용'; ws2['A1'].font = Font(size=18, bold=True)
    ws3 = wb.create_sheet('숨긴시트'); ws3['A1'] = '숨긴 시트 — 탭에 보이면 안 됨'; ws3.sheet_state = 'hidden'
    ws4 = wb.create_sheet('셋째(긴 이름의 시트 탭입니다)'); ws4['C3'] = '셋째'
    wb.save(f'{OUT}/xlsx/x07-multisheet-print.xlsx')


def x08():
    wb, ws = wb_new('견적서')
    for col, w in zip('ABCDEFG', [5, 26, 8, 8, 13, 15, 14]): ws.column_dimensions[col].width = w
    ws.merge_cells('A1:G1'); ws['A1'] = '견   적   서'; ws['A1'].font = Font(name='바탕', size=24, bold=True, underline='double'); ws['A1'].alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 45
    ws.merge_cells('A3:C3'); ws['A3'] = '2026년 11월 20일'; ws.merge_cells('A4:C4'); ws['A4'] = '(주)고객사 귀하'; ws['A4'].font = Font(size=13, bold=True, underline='single')
    ws.merge_cells('E3:F3'); ws['E3'] = '공급자: 향기랩 주식회사'; ws.merge_cells('E4:F4'); ws['E4'] = '사업자번호 123-45-67890'
    ws.merge_cells('A6:B6'); ws['A6'] = '합계 금액(부가세 포함)'; ws.merge_cells('C6:G6'); ws['C6'] = '=F20'; ws['C6'].number_format = '"일금 "#,##0"원정"'
    ws['C6'].font = Font(size=14, bold=True)
    for row in ws['A6:G6']:
        for c in row: c.border = Border(top=Side('medium'), bottom=Side('medium')); c.fill = PatternFill('solid', fgColor='FCE4D6')
    for i, h in enumerate(['No', '품명 및 규격', '단위', '수량', '단가', '공급가액', '비고'], 1):
        c = ws.cell(8, i, h); c.fill = PatternFill('solid', fgColor='44546A'); c.font = Font(bold=True, color='FFFFFF'); c.alignment = Alignment(horizontal='center'); c.border = BOX
    items = [('공조기 필터 교체(중성능, 592×592×292)', 'EA', 24, 85000), ('냉각탑 수처리 약품', 'kg', 200, 4500), ('배관 보온 공사', 'm', 120, 32000), ('제어반 점검 및 시운전', '식', 1, 1500000)]
    for r in range(9, 18):
        it = items[r - 9] if r - 9 < len(items) else None
        if it:
            ws.cell(r, 1, r - 8); ws.cell(r, 2, it[0]); ws.cell(r, 3, it[1]); ws.cell(r, 4, it[2]); ws.cell(r, 5, it[3])
        ws.cell(r, 6, f'=IF(D{r}="","",D{r}*E{r})')
        for c in range(1, 8):
            cell = ws.cell(r, c); cell.border = BOX
            if c in (5, 6): cell.number_format = '#,##0'
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='center')
    for r, (k, f) in enumerate([('공급가액', '=SUM(F9:F17)'), ('부가세', '=ROUND(F18*0.1,0)'), ('합계', '=F18+F19')], 18):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5); ws.cell(r, 1, k).alignment = Alignment(horizontal='right'); ws.cell(r, 6, f).number_format = '#,##0'
        for c in range(1, 8): ws.cell(r, c).border = BOX
    ws.cell(20, 6).font = Font(bold=True, color='C00000')
    img = XLImage(png(160, 64, 'logo')); img.width, img.height = 120, 48; ws.add_image(img, 'G2')
    ws.merge_cells('A22:G24'); ws['A22'] = '※ 유효기간: 견적일로부터 30일. 설치 일정은 협의 후 확정하며, 현장 여건에 따라 금액이 달라질 수 있습니다.'
    ws['A22'].alignment = Alignment(wrap_text=True, vertical='top')
    wb.save(f'{OUT}/xlsx/x08-korean-quotation.xlsx')


# ───────────────────────── PPTX ─────────────────────────
from pptx import Presentation
from pptx.util import Pt as PPt, Cm as PCm
from pptx.dml.color import RGBColor as PRGB
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE, MSO_ANCHOR
from pptx.enum.dml import MSO_THEME_COLOR, MSO_LINE_DASH_STYLE
A_NS = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'


def prs_new():
    p = Presentation(); p.slide_width, p.slide_height = PCm(33.867), PCm(19.05); return p


def p01():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[0]); s.shapes.title.text = '2026년 설비관리 사업 계획'; s.placeholders[1].text = '향기랩 설비관리팀 · 2026. 11.'
    s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = '추진 방향(글머리 수준)'
    tf = s.placeholders[1].text_frame; tf.text = '예방 정비 중심으로 전환'
    for lv, t in [(1, '고장 이력 분석 → 점검 주기 조정'), (2, '공조기·펌프 우선'), (1, '센서 기반 상태 감시'), (0, 'Energy saving 목표 10%'), (1, 'LED 교체, 인버터 도입'), (0, '안전 관리 강화')]:
        para = tf.add_paragraph(); para.text = t; para.level = lv
    s = p.slides.add_slide(p.slide_layouts[3]); s.shapes.title.text = '두 개 콘텐츠'
    s.placeholders[1].text_frame.text = '왼쪽: 현황'; s.placeholders[2].text_frame.text = '오른쪽: 개선안'
    s = p.slides.add_slide(p.slide_layouts[2]); s.shapes.title.text = '구역 머리글'; s.placeholders[1].text = '2부. 세부 계획'
    p.save(f'{OUT}/pptx/p01-placeholders-bullets.pptx')


def p02():
    p = prs_new()
    theme_part = [r.target_part for r in p.slide_master.part.rels.values() if r.reltype.endswith('/theme')][0]
    th = etree.fromstring(theme_part.blob)
    A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    for tag, ea_face, latin in (('majorFont', 'HY견고딕', 'Georgia'), ('minorFont', '맑은 고딕', 'Segoe UI')):
        th.find(f'.//{{{A}}}{tag}/{{{A}}}ea').set('typeface', ea_face)
        th.find(f'.//{{{A}}}{tag}/{{{A}}}latin').set('typeface', latin)
    th.find(f'.//{{{A}}}accent1')[0].set('val', 'ED7D31')
    th.find(f'.//{{{A}}}dk2')[0].set('val', '1F3864')
    theme_part._blob = etree.tostring(th, xml_declaration=True, encoding='UTF-8', standalone=True)
    s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = 'Theme fonts 테마 글꼴(제목: Georgia/HY견고딕)'
    s.placeholders[1].text_frame.text = '본문: Segoe UI / 맑은 고딕 — 직접 글꼴 지정 없음'
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PCm(20), PCm(12), PCm(10), PCm(4)); sh.text = '강조1 테마 색(주황)'
    sh2 = s.shapes.add_shape(MSO_SHAPE.OVAL, PCm(4), PCm(12), PCm(6), PCm(4)); sh2.fill.solid(); sh2.fill.fore_color.theme_color = MSO_THEME_COLOR.TEXT_2; sh2.text = '어두운2'
    p.save(f'{OUT}/pptx/p02-theme-fonts-colors.pptx')


def p03():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[1]); s.shapes.title.text = '자동 맞춤(글자 축소) — 넘치는 본문'
    tf = s.placeholders[1].text_frame; tf.text = KO
    for i in range(12): tf.add_paragraph().text = f'{i + 1}. ' + KO
    bp = tf._txBody.find(qn('a:bodyPr'))
    for ch in list(bp): bp.remove(ch)
    bp.append(parse_xml(f'<a:normAutofit {A_NS} fontScale="47500" lnSpcReduction="20000"/>'))
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = '도형에 맞게 크기 조정 / 넘침 / 세로 가운데'
    for i, (mode, txt) in enumerate([(MSO_AUTO_SIZE.SHAPE_TO_FIT_TEXT, '도형을 글에 맞춤 ' + KO), (MSO_AUTO_SIZE.NONE, '자동 맞춤 없음 — 넘침 ' + KO * 3), (None, '세로 가운데 짧은 글')]):
        tb = s.shapes.add_textbox(PCm(1 + i * 11), PCm(5), PCm(10), PCm(4)); tb.line.color.rgb = PRGB(0, 0, 0); tf = tb.text_frame; tf.word_wrap = True; tf.text = txt
        if mode is not None: tf.auto_size = mode
        else: tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p.save(f'{OUT}/pptx/p03-autofit.pptx')


def p04():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = '표: 병합·줄무늬·직접 채우기'
    t = s.shapes.add_table(6, 5, PCm(2), PCm(4.5), PCm(29), PCm(12)).table
    for i, h in enumerate(['구분', '항목', '1분기', '2분기', '비고']): t.cell(0, i).text = h
    rows = [['설비', '공조기', '1,200', '1,350', ''], ['', '펌프', '3,400', '0', '긴급'], ['전기', '수변전', '2,100', '2,300', ''], ['', '조명 LED 교체(본관 1~5층)', '5,600', '4,800', ''], ['합계', '', '12,300', '8,450', '']]
    for r, row in enumerate(rows, 1):
        for c, v in enumerate(row):
            if v: t.cell(r, c).text = v
    t.cell(1, 0).merge(t.cell(2, 0)); t.cell(3, 0).merge(t.cell(4, 0)); t.cell(5, 0).merge(t.cell(5, 1))
    t.cell(2, 4).fill.solid(); t.cell(2, 4).fill.fore_color.rgb = PRGB(0xF8, 0xCB, 0xAD)
    for c in range(2, 4):
        for r in range(1, 6): t.cell(r, c).text_frame.paragraphs[0].alignment = PP_ALIGN.RIGHT
    t.columns[1].width = PCm(10)
    p.save(f'{OUT}/pptx/p04-table.pptx')


def p05():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[5]); s.shapes.title.text = '차트: 막대·꺾은선·원형·누적 가로'
    cd = CategoryChartData(); cd.categories = ['1월', '2월', '3월', '4월']; cd.add_series('본관', (120, 135, 150, 142)); cd.add_series('별관', (80, 95, 70, 110))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PCm(1), PCm(4), PCm(15), PCm(7), cd).chart; ch.has_legend = True; ch.legend.position = XL_LEGEND_POSITION.BOTTOM
    ch.legend.include_in_layout = False; ch.plots[0].has_data_labels = True
    ch2 = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, PCm(17), PCm(4), PCm(15), PCm(7), cd).chart; ch2.has_title = True; ch2.chart_title.text_frame.text = '추세(꺾은선)'
    pd = CategoryChartData(); pd.categories = ['전기', '기계', '소방', '기타']; pd.add_series('비율', (0.45, 0.30, 0.15, 0.10))
    ch3 = s.shapes.add_chart(XL_CHART_TYPE.PIE, PCm(1), PCm(11.5), PCm(12), PCm(7), pd).chart; ch3.has_legend = True; ch3.plots[0].has_data_labels = True
    ch3.plots[0].data_labels.number_format = '0%'; ch3.plots[0].data_labels.number_format_is_linked = False
    s.shapes.add_chart(XL_CHART_TYPE.BAR_STACKED, PCm(17), PCm(11.5), PCm(15), PCm(7), cd)
    p.save(f'{OUT}/pptx/p05-charts.pptx')


def p06():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[6])
    a = s.shapes.add_picture(png(800, 600), PCm(1), PCm(1), PCm(10), PCm(7.5))
    b = s.shapes.add_picture(png(800, 600), PCm(12), PCm(1), PCm(8), PCm(7.5)); b.crop_left = 0.25; b.crop_right = 0.05; b.crop_top = 0.1
    c = s.shapes.add_picture(png(800, 600), PCm(22), PCm(1), PCm(10), PCm(7.5)); c.rotation = 15
    d_ = s.shapes.add_picture(png(800, 600), PCm(1), PCm(10), PCm(14), PCm(5))  # 비율 어긋남(늘림)
    s.shapes.add_picture(png(160, 64, 'logo'), PCm(28), PCm(16), PCm(4), PCm(1.6))
    for label, x, y in [('원본', 1, 8.7), ('자르기 L25%·R5%·T10%', 12, 8.7), ('15도 회전', 22, 8.7), ('가로로 늘림', 1, 15.2)]:
        s.shapes.add_textbox(PCm(x), PCm(y), PCm(10), PCm(1)).text_frame.text = label
    p.save(f'{OUT}/pptx/p06-pictures-crop-rotate.pptx')


def p07():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[6])
    g = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PCm(1), PCm(1), PCm(9), PCm(5)); g.text = '선형 그라데이션'
    g.fill.gradient(); g.fill.gradient_angle = 45; g.fill.gradient_stops[0].color.rgb = PRGB(0x1F, 0x4E, 0x79); g.fill.gradient_stops[1].color.rgb = PRGB(0x9D, 0xC3, 0xE6)
    sh = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PCm(12), PCm(1), PCm(9), PCm(5)); sh.text = '그림자(바깥쪽)'; sh.fill.solid(); sh.fill.fore_color.rgb = PRGB(255, 255, 255)
    sh.line.color.rgb = PRGB(0x44, 0x72, 0xC4); sh.text_frame.paragraphs[0].runs[0].font.color.rgb = PRGB(0, 0, 0)
    sh._element.spPr.append(parse_xml(f'<a:effectLst {A_NS}><a:outerShdw blurRad="76200" dist="76200" dir="2700000" algn="tl" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:outerShdw></a:effectLst>'))
    r = s.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, PCm(23), PCm(1.5), PCm(9), PCm(4)); r.rotation = 30; r.text = '30° 화살표'
    for i, geom in enumerate([MSO_SHAPE.OVAL, MSO_SHAPE.ISOSCELES_TRIANGLE, MSO_SHAPE.PENTAGON, MSO_SHAPE.STAR_5_POINT, MSO_SHAPE.CHEVRON, MSO_SHAPE.CAN, MSO_SHAPE.CLOUD, MSO_SHAPE.FLOWCHART_DECISION]):
        s.shapes.add_shape(geom, PCm(1 + i * 4), PCm(8), PCm(3.5), PCm(3)).text = str(i + 1)
    grp = s.shapes.add_group_shape()
    grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, PCm(2), PCm(13), PCm(6), PCm(3)).text = '그룹 A'
    grp.shapes.add_shape(MSO_SHAPE.RECTANGLE, PCm(12), PCm(13), PCm(6), PCm(3)).text = '그룹 B'
    cn = grp.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, PCm(8), PCm(14.5), PCm(12), PCm(14.5)); cn.line.width = PPt(2.5); cn.line.color.rgb = PRGB(0xC0, 0, 0)
    cn._element.spPr.find(qn('a:ln')).append(parse_xml(f'<a:tailEnd {A_NS} type="triangle"/>'))
    dash = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, PCm(21), PCm(13), PCm(10), PCm(4)); dash.fill.background(); dash.line.color.rgb = PRGB(0x70, 0x30, 0xA0); dash.line.width = PPt(3)
    dash.line.dash_style = MSO_LINE_DASH_STYLE.DASH; dash.text = '채우기 없음·점선 테두리'; dash.text_frame.paragraphs[0].runs[0].font.color.rgb = PRGB(0x70, 0x30, 0xA0)
    p.save(f'{OUT}/pptx/p07-shapes-effects-groups.pptx')


def p08():
    p = prs_new()
    s = p.slides.add_slide(p.slide_layouts[6])
    s.background.fill.solid(); s.background.fill.fore_color.rgb = PRGB(0xF2, 0xF2, 0xF2)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, p.slide_width, PCm(3)); band.fill.solid(); band.fill.fore_color.rgb = PRGB(0x1F, 0x38, 0x64); band.line.fill.background()
    tb = s.shapes.add_textbox(PCm(1.2), PCm(0.6), PCm(30), PCm(2)); r = tb.text_frame.paragraphs[0].add_run(); r.text = 'Ⅲ. 추진 일정 및 기대 효과'
    r.font.size = PPt(28); r.font.bold = True; r.font.color.rgb = PRGB(255, 255, 255); r.font.name = '맑은 고딕'
    tf = s.shapes.add_textbox(PCm(1.2), PCm(4), PCm(15), PCm(13)).text_frame; tf.word_wrap = True
    items = [('□ 1단계(12월)', 0, True), ('현황 조사 및 고장 이력 정리', 1, False), ('□ 2단계(2027년 1~3월)', 0, True), ('개선 공사: 공조기 4대, 펌프 6대 교체', 1, False),
             ('ㅇ 공사 중 정전 최소화(야간 작업)', 2, False), ('□ 기대 효과', 0, True), ('에너지 사용량 10% 절감(연 1.2억 원)', 1, False), ('고장 출동 건수 30% 감소', 1, False)]
    for i, (t, lv, b) in enumerate(items):
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph(); run = para.add_run(); run.text = ('    ' * lv) + ('- ' if lv else '') + t
        run.font.size = PPt(18 if lv == 0 else 15); run.font.bold = b; run.font.name = '맑은 고딕'; para.line_spacing = 1.3; para.space_after = PPt(6)
    cd = CategoryChartData(); cd.categories = ['현재', '1년 후', '2년 후']; cd.add_series('에너지(억 원)', (12.0, 10.8, 10.2))
    ch = s.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, PCm(17.5), PCm(4), PCm(15), PCm(9), cd).chart; ch.plots[0].has_data_labels = True
    ch.has_title = True; ch.chart_title.text_frame.text = '연간 에너지 비용'
    note = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, PCm(17.5), PCm(14), PCm(15), PCm(3.5)); note.fill.solid(); note.fill.fore_color.rgb = PRGB(0xFF, 0xF2, 0xCC)
    note.line.color.rgb = PRGB(0xBF, 0x90, 0x00); note.text_frame.text = '※ 금액은 2026년 단가 기준 추정치이며 실제 공사 입찰 결과에 따라 달라질 수 있음'
    run = note.text_frame.paragraphs[0].runs[0]; run.font.size = PPt(13); run.font.color.rgb = PRGB(0, 0, 0)
    s.notes_slide.notes_text_frame.text = '발표자 메모: 예산 질문 대비'
    p.save(f'{OUT}/pptx/p08-korean-business-slide.pptx')


if __name__ == '__main__':
    for f in (d01, d02, d03, d04, d05, d06, d07, d08, x01, x02, x03, x04, x05, x06, x07, x08, p01, p02, p03, p04, p05, p06, p07, p08):
        f(); print('ok', f.__name__)
