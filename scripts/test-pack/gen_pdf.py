# PDF test files: test-pdf.pdf (9 pages) and test-pdf-protected.pdf (password test1234).
import os
from tp_common import OUT, ASSETS, FONTS, Registry
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, legal, landscape
from reportlab.lib.colors import HexColor, Color, black, white, red, blue, green
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from pypdf import PdfReader, PdfWriter

R = Registry('pdf')
FILE = 'test-pdf.pdf'
pdfmetrics.registerFont(TTFont('ArialEmbedded', os.path.join(FONTS, 'arial.ttf')))
pdfmetrics.registerFont(TTFont('CalibriEmbedded', os.path.join(FONTS, 'calibri.ttf')))
for cid in ('HeiseiMin-W3', 'STSong-Light', 'HYSMyeongJo-Medium'):
    pdfmetrics.registerFont(UnicodeCIDFont(cid))

NAVY = HexColor('#1F4E79')
c = canvas.Canvas(os.path.join(OUT, FILE), pagesize=letter)
c.setTitle('Docs Viewer test pack - PDF'); c.setAuthor('Docs Viewer test pack')
page = [0]

def start(title, size=letter):
    page[0] += 1
    c.setPageSize(size)
    w, h = size
    c.setFillColor(NAVY); c.rect(0, h - 60, w, 60, fill=1, stroke=0)
    c.setFillColor(white); c.setFont('Helvetica-Bold', 20); c.drawString(36, h - 40, f'Page {page[0]} · {title}')
    c.bookmarkPage(f'p{page[0]}'); c.addOutlineEntry(f'Page {page[0]}: {title}', f'p{page[0]}', level=0)
    return w, h

def tests_box(w, items, y=150):
    """Tests on this page, printed at the bottom."""
    c.setFillColor(HexColor('#F2F2F2')); c.rect(30, 30, w - 60, y - 20, fill=1, stroke=0)
    c.setFillColor(NAVY); c.setFont('Helvetica-Bold', 11); c.drawString(40, y - 10, 'Tests on this page')
    yy = y - 28
    for what, expected, *status in items:
        tid = f'PDF-{len(R.tests) + 1:02d}'
        R.add(tid, FILE, f'Page {page[0]}', what, expected, status[0] if status else '')
        c.setFillColor(NAVY); c.setFont('Helvetica-Bold', 8.5); c.drawString(40, yy, f'{tid} · {what}')
        c.setFillColor(HexColor('#595959')); c.setFont('Helvetica-Oblique', 8); c.drawString(40, yy - 10, f'Expected: {expected}'[:150])
        yy -= 24

# 1. Text and fonts
w, h = start('Text and fonts')
y = h - 100
for font, label in [('Helvetica', 'Helvetica (standard PDF font)'), ('Times-Roman', 'Times (standard PDF font)'), ('Courier', 'Courier (standard PDF font)'),
                    ('ArialEmbedded', 'Arial (embedded font)'), ('CalibriEmbedded', 'Calibri (embedded font)')]:
    c.setFillColor(black); c.setFont(font, 16); c.drawString(40, y, f'{label}: The quick brown fox 0123456789'); y -= 30
for size in (8, 12, 18, 28):
    c.setFont('Helvetica', size); c.drawString(40, y, f'{size} pt text'); y -= size + 12
for col, name in ((red, 'Red'), (blue, 'Blue'), (green, 'Green')):
    c.setFillColor(col); c.setFont('Helvetica-Bold', 16); c.drawString(40, y, f'{name} text'); y -= 26
c.setFillColor(black)
c.setFont('HeiseiMin-W3', 16); c.drawString(40, y, '日本語のテキスト (Japanese)'); y -= 26
c.setFont('STSong-Light', 16); c.drawString(40, y, '中文文本 (Chinese)'); y -= 26
c.setFont('HYSMyeongJo-Medium', 16); c.drawString(40, y, '한국어 텍스트 (Korean)'); y -= 30
c.setFont('Helvetica', 14); c.drawString(40, y, 'Search test word: FINDME   (and again here: FINDME)')
tests_box(w, [('Standard fonts', 'Helvetica, Times, Courier lines in those styles'), ('Embedded fonts', 'Arial and Calibri lines look like those fonts'),
              ('Sizes and colors', '8 to 28 pt lines; red, blue, green text'), ('Japanese, Chinese, Korean', 'Three lines of Asian characters (no empty boxes)'),
              ('Search (Ctrl+F)', 'Searching FINDME finds 2 matches on page 1'), ('Select and copy text', 'Text can be selected and copied')], y=175)
c.showPage()

# 2. Vector drawings
w, h = start('Vector drawings')
c.setStrokeColor(NAVY); c.setLineWidth(3); c.line(40, h - 110, w - 40, h - 110)
c.setDash(6, 4); c.line(40, h - 130, w - 40, h - 130); c.setDash()
c.setFillColor(HexColor('#4472C4')); c.rect(40, h - 300, 150, 120, fill=1, stroke=0)
c.setFillColor(HexColor('#ED7D31')); c.circle(300, h - 240, 60, fill=1, stroke=0)
c.setStrokeColor(HexColor('#70AD47')); c.setLineWidth(4); c.roundRect(400, h - 300, 160, 120, 20, fill=0, stroke=1)
p = c.beginPath(); p.moveTo(40, h - 340); p.lineTo(120, h - 460); p.lineTo(200, h - 340); p.close()
c.setFillColor(HexColor('#FFC000')); c.drawPath(p, fill=1, stroke=0)
c.linearGradient(260, h - 460, 560, h - 340, (white, NAVY), extend=False)
c.setFillColor(black); c.setFont('Helvetica', 11); c.drawString(260, h - 475, '^ gradient box: white (left) to dark blue (right)')
tests_box(w, [('Lines', 'Solid and dashed dark blue lines across the page'), ('Filled shapes', 'Blue rectangle, orange circle, yellow triangle'),
              ('Outline shape', 'Green rounded rectangle outline'), ('Gradient', 'Box fades from white to dark blue')])
c.showPage()

# 3. Images
w, h = start('Images')
c.drawImage(os.path.join(ASSETS, 'picture.jpg'), 40, h - 400, width=420, height=280)
c.drawImage(os.path.join(ASSETS, 'picture2.png'), 40, h - 600, width=240, height=160)
tests_box(w, [('JPEG image', 'The "Embedded picture" test chart, colors matching the bar names'), ('PNG image', 'Purple "Picture 2" below it'),
              ('Zoom', 'Zooming in keeps the image and text sharp')])
c.showPage()

# 4. Links and bookmarks
w, h = start('Links and bookmarks')
c.setFillColor(HexColor('#0563C1')); c.setFont('Helvetica', 16)
c.drawString(40, h - 120, 'External link: https://example.com'); c.linkURL('https://example.com', (40, h - 125, 330, h - 105))
c.drawString(40, h - 160, 'Internal link: go to page 1'); c.linkRect('', 'p1', (40, h - 165, 260, h - 145))
tests_box(w, [('External link', 'Clicking "https://example.com" opens the website (new tab)', '⚠️'),
              ('Internal link', 'Clicking "go to page 1" jumps to page 1', '⚠️'),
              ('Bookmarks / outline', 'The file has 9 bookmarks (one per page), if the viewer shows bookmarks', '❌')])
c.showPage()

# 5. Form fields
w, h = start('Form fields (fillable PDF)')
f = c.acroForm
c.setFillColor(black); c.setFont('Helvetica', 12)
c.drawString(40, h - 110, 'Name:'); f.textfield(name='name', value='Jane Doe', x=120, y=h - 120, width=250, height=22)
c.drawString(40, h - 150, 'Agree:'); f.checkbox(name='agree', x=120, y=h - 155, size=18, checked=True)
c.drawString(40, h - 190, 'Account:'); f.choice(name='acct', value='Checking', options=['Checking', 'Savings'], x=120, y=h - 200, width=150, height=22)
c.drawString(40, h - 230, 'Type:'); f.radio(name='type', value='a', selected=True, x=120, y=h - 235, size=16); c.drawString(140, h - 230, 'Personal')
f.radio(name='type', value='b', selected=False, x=230, y=h - 235, size=16); c.drawString(250, h - 230, 'Business')
tests_box(w, [('Text field', 'Box showing "Jane Doe"', '⚠️'), ('Checkbox', 'Checked box next to Agree', '⚠️'),
              ('Drop-down', 'Box showing "Checking"', '⚠️'), ('Radio buttons', '"Personal" selected, "Business" not', '⚠️')])
c.showPage()

# 6-8. Page sizes and orientation
w, h = start('A4 page size', A4)
tests_box(w, [('A4 page', 'This page is A4 (slightly narrower and taller than Letter pages 1-5)')])
c.showPage()
w, h = start('Landscape page', landscape(letter))
tests_box(w, [('Landscape page', 'This page is wider than it is tall')])
c.showPage()
w, h = start('Legal page size', legal)
tests_box(w, [('Legal page', 'This page is taller than the Letter pages (8.5 x 14 in)')])
c.showPage()

# 9. Scan-only page (image, no text)
w, h = start('Scanned page (image only)')
c.drawImage(os.path.join(ASSETS, 'picture.jpg'), 40, h - 520, width=530, height=353)
tests_box(w, [('Scanned page without text', 'Shows the picture; searching for words on this page finds nothing (it is an image)')])
c.showPage()
c.save()

# Rotate page 7 by 90 degrees (stored rotation), add the rotation test
reader = PdfReader(os.path.join(OUT, FILE)); writer = PdfWriter()
for i, pg in enumerate(reader.pages):
    if i == 3: pass
    writer.add_page(pg)
writer.pages[2].rotate(90)
R.add(f'PDF-{len(R.tests) + 1:02d}', FILE, 'Page 3', 'Page stored with 90° rotation', 'Page 3 (Images) is shown turned sideways (rotated 90° clockwise), as in Adobe Reader')
R.add(f'PDF-{len(R.tests) + 1:02d}', FILE, 'All pages', 'Page count and thumbnails', 'Page bar shows 9 pages; thumbnails show different page sizes')
with open(os.path.join(OUT, FILE), 'wb') as fh:
    writer.write(fh)

# Password-protected copy
PFILE = 'test-pdf-protected.pdf'
pw = PdfWriter(clone_from=os.path.join(OUT, FILE)); pw.encrypt(user_password='test1234', owner_password='owner-test1234', algorithm='AES-256')
with open(os.path.join(OUT, PFILE), 'wb') as fh:
    pw.write(fh)
R.add('PDF-P1', PFILE, 'On open', 'Password prompt', 'Asks for a password before showing anything')
R.add('PDF-P2', PFILE, 'Password box', 'Wrong password', 'Typing "wrong" shows "Incorrect password"')
R.add('PDF-P3', PFILE, 'Password box', 'Correct password', 'Typing test1234 opens the same 9 pages as test-pdf.pdf')
R.save()
