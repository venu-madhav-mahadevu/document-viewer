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
    ['onboarding-guide.doc', '.doc Word 97-2003 (simplified preview)'],
  ] },
  { title: 'Excel', files: [
    ['regional-sales.xlsx', '.xlsx, merged cells, 2 sheets'],
    ['branches-macro.xlsm', '.xlsm macro-enabled'],
    ['branches-legacy.xls', '.xls Excel 97-2003'],
    ['accounts.csv', '.csv with quoted commas'],
  ] },
  { title: 'PowerPoint', files: [
    ['strategy-deck.pptx', '.pptx'],
    ['strategy-deck.ppt', '.ppt PowerPoint 97-2003 (simplified preview)'],
  ] },
  { title: 'Images', files: [
    ['photo.jpg', '.jpg'],
    ['photo-chart.jpeg', '.jpeg'],
    ['scanned-contract.tif', '.tif 3-page scan (fax G4 compression)'],
    ['color-chart.tiff', '.tiff color (LZW)'],
  ] },
  { title: 'Text', files: [
    ['readme.txt', '.txt with Unicode'],
  ] },
  // Large-file tests. large/large-scan-60-pages.tif is kept in the folder but not
  // listed: decoding 60 full-resolution pages at once freezes the tab.
  { title: 'Large files', files: [
    ['large/large-report-49mb.pdf', '49 MB, 100 pages'],
    ['large/large-transactions-50mb.xlsx', '50 MB, 370,000 rows'],
  ] },
];
