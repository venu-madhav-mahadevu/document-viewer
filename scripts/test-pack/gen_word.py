# Word test document. usage: python gen_word.py <PREFIX> <output.docx>
#   PREFIX = DOCX | DOCM | DOC  (the .docm and .doc copies are made from this file afterwards)
import os, sys, copy
from tp_common import OUT, ASSETS, Registry
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import parse_xml
from lxml import etree

PREFIX, OUTFILE = sys.argv[1], sys.argv[2]
FILE = {'DOCX': 'test-word.docx', 'DOCM': 'test-word.docm', 'DOC': 'test-word.doc'}[PREFIX]
R = Registry(f'word-{PREFIX.lower()}')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"')

doc = Document()
st = doc.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(11)
n = [0]

def label(where, what, expected, status=''):
    """Gray test label paragraph: ID, what is tested, expected result."""
    n[0] += 1
    tid = f'{PREFIX}-{n[0]:02d}'
    R.add(tid, FILE, where, what, expected, status)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(2)
    r = p.add_run(f'{tid} · {what}'); r.bold = True; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
    r2 = p.add_run(f'   Expected: {expected}'); r2.font.size = Pt(9); r2.font.color.rgb = RGBColor(0x59, 0x59, 0x59); r2.italic = True
    return tid

def page_title(text):
    h = doc.add_heading(text, level=1)
    return h

# ---------------------------------------------------------------- page 1: text
t = doc.add_paragraph(); r = t.add_run(f'{FILE} — Word test document'); r.bold = True; r.font.size = Pt(22); r.font.color.rgb = RGBColor(0x1F, 0x4E, 0x79)
doc.add_paragraph('Each test below has an ID (for the checklist), what is tested, and the expected result. '
                  'Compare what you see with "Expected". Open the same file in Microsoft Word to compare if unsure.')
if PREFIX == 'DOCM':
    label('Page 1', 'Macro-enabled document (.docm)', 'Opens normally. No macro runs, no security warning')
if PREFIX == 'DOC':
    label('Page 1', 'Word 97-2003 binary document (.doc)', 'Opens normally with the formatting below')

page_title('1. Fonts and text')
label('Page 1, section 1', 'Font families', 'Each line is shown in the font it names (Calibri, Cambria, Arial, Times New Roman, Courier New, Georgia, Verdana)')
for f in ['Calibri', 'Cambria', 'Arial', 'Times New Roman', 'Courier New', 'Georgia', 'Verdana']:
    p = doc.add_paragraph(); r = p.add_run(f'This line is {f}. The quick brown fox jumps over the lazy dog 0123456789'); r.font.name = f
    r._element.rPr.rFonts.set(qn('w:eastAsia'), f)

label('Page 1', 'Font sizes', 'Five lines growing from 8 pt to 28 pt')
for s in [8, 11, 14, 20, 28]:
    doc.add_paragraph().add_run(f'This text is {s} pt').font.size = Pt(s)

label('Page 1', 'Text styles', 'Each word looks like its name: bold, italic, underline, double underline, strikethrough, x² superscript, H₂O subscript, SMALL CAPS, ALL CAPS')
p = doc.add_paragraph()
for text, fmt in [('Bold', 'bold'), ('Italic', 'italic'), ('Underline', 'underline'), ('Double underline', 'double'),
                  ('Strikethrough', 'strike'), ('Small caps', 'small_caps'), ('all caps', 'all_caps')]:
    r = p.add_run(text)
    if fmt == 'double': r.font.underline = 3  # WD_UNDERLINE.DOUBLE
    else: setattr(r.font, fmt, True)
    p.add_run('   ')
p.add_run('x'); sup = p.add_run('2'); sup.font.superscript = True; p.add_run(' superscript   H'); sub = p.add_run('2'); sub.font.subscript = True; p.add_run('O subscript')

label('Page 1', 'Text colors and highlight', 'Red, blue, green and orange text; yellow and bright green highlight; a paragraph with light blue background')
p = doc.add_paragraph()
for text, rgb in [('Red text', 'C00000'), ('Blue text', '0070C0'), ('Green text', '00B050'), ('Orange text', 'ED7D31')]:
    p.add_run(text + '   ').font.color.rgb = RGBColor.from_string(rgb)
r = p.add_run('Yellow highlight'); r.font.highlight_color = WD_COLOR_INDEX.YELLOW
p.add_run('   ')
r = p.add_run('Green highlight'); r.font.highlight_color = WD_COLOR_INDEX.BRIGHT_GREEN
p = doc.add_paragraph('This whole paragraph has a light blue background (paragraph shading).')
p._p.get_or_add_pPr().append(parse_xml(f'<w:shd {NS} w:val="clear" w:color="auto" w:fill="DDEBF7"/>'))

label('Page 1', 'Alignment', 'Lines aligned left, centered, right; the long paragraph is justified (straight right edge)')
for text, al in [('Left aligned', WD_ALIGN_PARAGRAPH.LEFT), ('Centered', WD_ALIGN_PARAGRAPH.CENTER), ('Right aligned', WD_ALIGN_PARAGRAPH.RIGHT)]:
    doc.add_paragraph(text).alignment = al
p = doc.add_paragraph('Justified: ' + 'This paragraph is justified so both the left and the right edges are straight lines. ' * 3)
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

label('Page 1', 'Indents and line spacing', 'Indented paragraph (1 inch), first-line indent, hanging indent; then 1.0, 1.5 and 2.0 line spacing')
p = doc.add_paragraph('Indented 1 inch from the left. ' * 3); p.paragraph_format.left_indent = Inches(1)
p = doc.add_paragraph('First line indented half an inch. ' * 3); p.paragraph_format.first_line_indent = Inches(0.5)
p = doc.add_paragraph('Hanging indent: the first line starts left, the next lines are indented. ' * 2)
p.paragraph_format.left_indent = Inches(0.5); p.paragraph_format.first_line_indent = Inches(-0.5)
for sp in [1.0, 1.5, 2.0]:
    p = doc.add_paragraph(f'Line spacing {sp}. ' * 8); p.paragraph_format.line_spacing = sp

label('Page 1', 'Headings and styles', 'Title style, then Heading 1, 2, 3 in decreasing size and in the theme color')
doc.add_paragraph('Title style', style='Title')
for lv in [1, 2, 3]:
    doc.add_heading(f'Heading {lv}', level=lv)

# ---------------------------------------------------------------- page 2: lists
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
page_title('2. Bullets and numbering')
numbering = doc.part.numbering_part.element

def add_list(levels):
    """levels: [(fmt, text, font), ...] per level -> numId."""
    absid = 100 + len(numbering.findall(qn('w:abstractNum')))
    lv = ''.join(
        f'<w:lvl w:ilvl="{i}"><w:start w:val="1"/><w:numFmt w:val="{fmt}"/><w:lvlText w:val="{text}"/><w:lvlJc w:val="left"/>'
        f'<w:pPr><w:ind w:left="{720 * (i + 1)}" w:hanging="360"/></w:pPr>'
        + (f'<w:rPr><w:rFonts w:ascii="{fnt}" w:hAnsi="{fnt}" w:hint="default"/></w:rPr>' if fnt else '') + '</w:lvl>'
        for i, (fmt, text, fnt) in enumerate(levels))
    absn = parse_xml(f'<w:abstractNum {NS} w:abstractNumId="{absid}"><w:multiLevelType w:val="hybridMultilevel"/>{lv}</w:abstractNum>')
    first_num = numbering.find(qn('w:num'))
    if first_num is not None: first_num.addprevious(absn)
    else: numbering.append(absn)
    numid = 100 + len(numbering.findall(qn('w:num')))
    numbering.append(parse_xml(f'<w:num {NS} w:numId="{numid}"><w:abstractNumId w:val="{absid}"/></w:num>'))
    return numid

def list_item(text, numid, level=0):
    p = doc.add_paragraph(text)
    p._p.get_or_add_pPr().append(parse_xml(f'<w:numPr {NS}><w:ilvl w:val="{level}"/><w:numId w:val="{numid}"/></w:numPr>'))
    return p

for what, expected, levels in [
    ('Round bullets (Symbol font)', 'Round • bullets', [('bullet', '', 'Symbol')]),
    ('Square bullets (Wingdings §)', 'Small square ▪ bullets — NOT the § sign and NOT empty boxes', [('bullet', '', 'Wingdings')]),
    ('Arrow bullets (Wingdings Ø)', 'Arrow-head ➢ bullets — NOT Ø', [('bullet', '', 'Wingdings')]),
    ('Check-mark bullets (Wingdings ü)', 'Check ✓ bullets — NOT ü', [('bullet', '', 'Wingdings')]),
    ('Hollow circle bullets (Courier New o)', 'Hollow o bullets', [('bullet', 'o', 'Courier New')]),
    ('Dash bullets', 'Dash – bullets', [('bullet', '–', None)]),
]:
    label('Page 2', what, expected)
    nid = add_list(levels)
    for i in range(3): list_item(f'{what.split(" (")[0]} item {i + 1}', nid)

label('Page 2', 'Numbered list', 'Numbered 1. 2. 3.')
nid = add_list([('decimal', '%1.', None)])
for i in range(3): list_item(f'Numbered item {i + 1}', nid)

label('Page 2', 'Multi-level list (1 / a / i / bullet)', 'Level 1 numbered 1., level 2 a), level 3 i., level 4 ▪ bullet; each level indented more')
nid = add_list([('decimal', '%1.', None), ('lowerLetter', '%2)', None), ('lowerRoman', '%3.', None), ('bullet', '', 'Wingdings')])
for text, lvl in [('First level', 0), ('Second level', 1), ('Third level', 2), ('Fourth level (square bullet)', 3), ('Second level again', 1), ('First level again', 0)]:
    list_item(text, nid, lvl)

# ---------------------------------------------------------------- page 3: tables
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
page_title('3. Tables')
label('Page 3', 'Table borders, header shading, column widths', 'Dark blue header row with white bold text; all cells have borders; numbers right-aligned')
tbl = doc.add_table(rows=1, cols=4); tbl.style = 'Table Grid'; tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
for i, h in enumerate(['Region', 'Q1', 'Q2', 'Total']):
    c = tbl.rows[0].cells[i]; c.text = ''
    r = c.paragraphs[0].add_run(h); r.bold = True; r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    c._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {NS} w:val="clear" w:color="auto" w:fill="1F4E79"/>'))
for row in [('West', '1,200', '1,350', '2,550'), ('East', '1,300', '1,400', '2,700'), ('North', '1,400', '1,450', '2,850')]:
    cells = tbl.add_row().cells
    for i, v in enumerate(row):
        cells[i].text = v
        if i: cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT

label('Page 3', 'Merged cells and cell colors', 'Top row is one merged cell; "Merged down" spans two rows; green, yellow and red cell backgrounds')
t2 = doc.add_table(rows=3, cols=3); t2.style = 'Table Grid'
a = t2.cell(0, 0).merge(t2.cell(0, 2)); a.text = 'Merged across three columns'; a.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
b = t2.cell(1, 0).merge(t2.cell(2, 0)); b.text = 'Merged down'
for (r_, c_, txt, fill) in [(1, 1, 'Green', 'C6EFCE'), (1, 2, 'Yellow', 'FFEB9C'), (2, 1, 'Red', 'FFC7CE'), (2, 2, 'Plain', None)]:
    cell = t2.cell(r_, c_); cell.text = txt
    if fill: cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {NS} w:val="clear" w:color="auto" w:fill="{fill}"/>'))

# ---------------------------------------------------------------- page 4: pictures and shapes
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
page_title('4. Pictures, shapes and arrows')
label('Page 4', 'Inline picture', 'The "Embedded picture" test chart is shown at about 4 inches wide')
doc.add_picture(os.path.join(ASSETS, 'picture.png'), width=Inches(4))

def anchor_xml(inner_graphic, x_in, y_in, w_in, h_in, doc_id, behind=False, wrap='<wp:wrapNone/>'):
    emu = lambda v: int(v * 914400)
    return (f'<w:r {NS}><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{251658240 + doc_id}" '
            f'behindDoc="{1 if behind else 0}" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="column"><wp:posOffset>{emu(x_in)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{emu(y_in)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(w_in)}" cy="{emu(h_in)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
            f'<wp:docPr id="{doc_id}" name="Shape {doc_id}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">{inner_graphic}</a:graphicData></a:graphic>'
            f'</wp:anchor></w:drawing></w:r>')

def shape(prst, fill, line, w_in, h_in, text='', rot=0, line_w=19050, no_fill=False, arrow_end=False):
    emu = lambda v: int(v * 914400)
    fill_xml = '<a:noFill/>' if no_fill else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    tail = '<a:tailEnd type="triangle"/>' if arrow_end else ''
    body = (f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            if text else '')
    return (f'<wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm rot="{rot * 60000}"><a:off x="0" y="0"/><a:ext cx="{emu(w_in)}" cy="{emu(h_in)}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill>{tail}</a:ln></wps:spPr>'
            f'{body}<wps:bodyPr anchor="ctr"/></wps:wsp>')

label('Page 4', 'Shapes and arrows over a picture', 'Over the picture below: a red right-arrow at the left pointing at the gray ramp, a yellow rounded box "Look here" at the top right, '
      'a green circle outline at the right end of the gray ramp, and a blue line with an arrow head below the fine lines', '❌')
host = doc.add_paragraph()
host.paragraph_format.space_after = Pt(0)
host.add_run().add_picture(os.path.join(ASSETS, 'picture.png'), width=Inches(5))
shapes = [
    (shape('rightArrow', 'FF0000', 'C00000', 1.2, 0.45), 0.1, 1.35, 1.2, 0.45),
    (shape('roundRect', 'FFC000', 'BF9000', 1.5, 0.5, 'Look here'), 3.3, 0.2, 1.5, 0.5),
    (shape('ellipse', 'FFFFFF', '00B050', 0.9, 0.9, no_fill=True, line_w=38100), 3.9, 1.4, 0.9, 0.9),
    (shape('line', 'FFFFFF', '0070C0', 1.6, 0.6, line_w=28575, arrow_end=True), 1.6, 2.4, 1.6, 0.6),
]
for i, (g, x, y, w, h) in enumerate(shapes):
    host._p.append(parse_xml(anchor_xml(g, x, y, w, h, 200 + i)))

label('Page 4', 'Text box', 'A light gray box with a dark border containing two lines of text, to the right of this paragraph', '❌')
p = doc.add_paragraph('This paragraph has a text box placed next to it. ' * 3)
tb = (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(2.2 * 914400)}" cy="{int(0.8 * 914400)}"/></a:xfrm>'
      f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F2F2F2"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="404040"/></a:solidFill></a:ln></wps:spPr>'
      f'<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Text box title</w:t></w:r></w:p><w:p><w:r><w:t>Second line inside the box</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
      f'<wps:bodyPr/></wps:wsp>')
p._p.append(parse_xml(anchor_xml(tb, 4.2, 0, 2.2, 0.8, 210, wrap='<wp:wrapSquare wrapText="bothSides"/>')))

label('Page 4', 'Floating picture with text wrap', 'The small purple "Picture 2" sits on the right; this text wraps around its left side', '⚠️')
p = doc.add_paragraph('Text wraps around the floating picture. ' * 12)
rel_id, _ = doc.part.get_or_add_image(os.path.join(ASSETS, 'picture2.png'))
pic = (f'<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="picture2.png"/><pic:cNvPicPr/></pic:nvPicPr>'
       f'<pic:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
       f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(2 * 914400)}" cy="{int(1.333 * 914400)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>')
anch = anchor_xml(pic, 4.3, 0, 2, 1.333, 220, wrap='<wp:wrapSquare wrapText="bothSides"/>').replace(
    'http://schemas.microsoft.com/office/word/2010/wordprocessingShape', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
p._p.insert(0, parse_xml(anch))

# ---------------------------------------------------------------- page 5: links, equation
doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
page_title('5. Links and equation')
label('Page 5', 'Hyperlink', 'Blue underlined link "Open example.com"; clicking it opens https://example.com in a new tab')
p = doc.add_paragraph()
rid = doc.part.relate_to('https://example.com', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', is_external=True)
p._p.append(parse_xml(f'<w:hyperlink {NS} r:id="{rid}"><w:r><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr><w:t>Open example.com</w:t></w:r></w:hyperlink>'))

label('Page 5', 'Equation (math)', 'The formula E = mc² and a fraction a over b, shown as math (not plain text)', '❌')
p = doc.add_paragraph()
p._p.append(parse_xml(f'<m:oMath {NS}><m:r><m:t>E=m</m:t></m:r><m:sSup><m:e><m:r><m:t>c</m:t></m:r></m:e><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>'
                      f'<m:r><m:t>,  </m:t></m:r><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>'))

label('Page 5', 'Page break', 'This test is at the top of page 5 (pages 1-4 end with page breaks)')
label('Every page', 'Header and footer', 'Header "Docs Viewer test pack" at the top of every page; footer "Page N" with the right number', '')

# header / footer with page number field
sec = doc.sections[0]
hp = sec.header.paragraphs[0]; hr = hp.add_run('Docs Viewer test pack — header'); hr.font.size = Pt(9); hr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
fp = sec.footer.paragraphs[0]; fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
fp.add_run('Page ')
fp._p.append(parse_xml(f'<w:fldSimple {NS} w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>'))

# ---------------------------------------------------------------- section 2: two columns
new = doc.add_section(WD_SECTION.NEW_PAGE)
page_title('6. Two columns')
label('Page 6', 'Two-column layout', 'Text flows in two columns side by side (like a newsletter)', '⚠️')
for i in range(6):
    doc.add_paragraph(f'Column text paragraph {i + 1}. ' + 'Text continues in two columns. ' * 6)
new._sectPr.find(qn('w:cols')).set(qn('w:num'), '2')  # one <w:cols> per section (Word rejects two)

# ---------------------------------------------------------------- section 3: landscape
land = doc.add_section(WD_SECTION.NEW_PAGE)
land._sectPr.find(qn('w:cols')).attrib.pop(qn('w:num'), None)  # back to one column (add_section copies the last one)
land.orientation = WD_ORIENT.LANDSCAPE
land.page_width, land.page_height = land.page_height, land.page_width
page_title('7. Landscape page')
label('Page 7', 'Landscape page in a portrait document', 'This page is wider than it is tall (landscape); the pages before it are portrait')
doc.add_paragraph('Landscape page content. ' * 20)

doc.save(OUTFILE)
R.save()
