// Demo files in public/samples, grouped for the home-screen gallery.
export const SAMPLE_GROUPS = [
  { title: 'PDF', files: [
    ['quarterly-report.pdf', '12 pages, selectable text'],
    ['mixed-page-sizes.pdf', 'Letter + A4, portrait + landscape'],
    ['protected-statement.pdf', 'Password: demo1234'],
  ] },
  { title: 'Word', files: [
    ['onboarding-guide.docx', '.docx'],
    ['onboarding-macro.docm', '.docm macro-enabled (macros never run)'],
    ['onboarding-guide.doc', '.doc Word 97-2003'],
  ] },
  { title: 'Excel', files: [
    ['regional-sales.xlsx', '.xlsx colorful: shading, merged cells, totals, 2 sheets'],
    ['branches-macro.xlsm', '.xlsm macro-enabled, color-coded ratios'],
    ['branches-legacy.xls', '.xls Excel 97-2003 with cell colors'],
    ['accounts.csv', '.csv with quoted commas'],
    ['formatted-report.xlsx', '.xlsx formats and colors: currency, %, dates, fills, fonts, borders (3 sheets)'],
    ['formatted-report.xls', '.xls Excel 97-2003 copy of the formats and colors file'],
  ] },
  { title: 'PowerPoint', files: [
    ['strategy-deck.pptx', '.pptx'],
    ['strategy-deck.ppt', '.ppt PowerPoint 97-2003'],
  ] },
  { title: 'Images', files: [
    ['photo.jpg', '.jpg'],
    ['photo-chart.jpeg', '.jpeg'],
    ['scanned-contract.tif', '.tif 3-page scan (fax G4 compression)'],
    ['color-chart.tiff', '.tiff color (LZW)'],
    ['color-scan-jpeg.tif', '.tif color scan (JPEG compression)'],
  ] },
  { title: 'Text', files: [
    ['readme.txt', '.txt with Unicode'],
  ] },
  { title: 'Large files', files: [
    ['large/large-report-49mb.pdf', '49 MB, 100 pages'],
    ['large/large-transactions-50mb.xlsx', '50 MB, 370,000 rows'],
    ['large/large-scan-60-pages.tif', '36 MB, 60-page scan'],
  ] },
];
