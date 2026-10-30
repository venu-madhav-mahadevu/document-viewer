# CSV test file: quoting, commas, line breaks, accents, Asian text, numbers, dates.
import csv, io, os
from tp_common import OUT, Registry

R = Registry('csv')
FILE = 'test-data.csv'
rows = [['Test ID', 'What is tested', 'Test value', 'Expected (what you should see in "Test value")']]
def t(what, value, expected, status=''):
    tid = f'CSV-{len(rows):02d}'
    R.add(tid, FILE, f'Row {len(rows) + 1} (Test ID {tid})', what, expected, status)
    rows.append([tid, what, value, expected])

t('Plain text', 'Hello world', 'Hello world')
t('Comma inside quotes', 'Smith, John', 'Smith, John (one cell, the comma is part of the text)')
t('Quotes inside quotes', 'He said "hello"', 'He said "hello" (with the quote marks)')
t('Line break inside a cell', 'First line\nSecond line', 'Two lines inside one cell')
t('Accented characters (UTF-8)', 'café naïve Zürich São Paulo', 'café naïve Zürich São Paulo (no garbled characters like Ã©)')
t('Asian text', '日本語 中文 한국어', '日本語 中文 한국어')
t('Emoji', 'Done ✅ Thanks 🙂', 'Done ✅ Thanks 🙂')
t('Leading zeros', '00042', '00042 (zeros kept, not 42)', '⚠️')
t('Long number (account / card)', '1234567812345678', '1234567812345678 (all digits, not 1.23E+15)', '⚠️')
t('Decimal number', '1234.56', '1234.56')
t('Negative number', '-75.25', '-75.25')
t('Currency text', '$1,234.50', '$1,234.50')
t('Percent text', '45%', '45%')
t('ISO date', '2026-10-21', '2026-10-21 (or the same date in another format)')
t('US date', '10/21/2026', '10/21/2026')
t('Empty value', '', 'Empty cell')
t('Spaces at start and end', '   padded   ', 'padded (with spaces around it)')
t('Very long text', 'Lorem ipsum ' * 30, 'Long text; the cell does not break the table layout')
t('Formula-looking text (safety)', '=1+1', '=1+1 shown as text (NOT calculated as 2)')

buf = io.StringIO()
csv.writer(buf, lineterminator='\r\n').writerows(rows)
# UTF-8 with a byte-order mark, as Excel saves "CSV UTF-8"
with open(os.path.join(OUT, FILE), 'w', encoding='utf-8-sig', newline='') as f:
    f.write(buf.getvalue())
R.add('CSV-20', FILE, 'Whole file', 'UTF-8 file with byte-order mark (Excel "CSV UTF-8")', 'First header cell reads "Test ID" (no strange characters before it)')
R.add('CSV-21', FILE, 'Whole file', 'Table view', 'Shown as a table with 4 columns and 20 rows of data')
R.save()
