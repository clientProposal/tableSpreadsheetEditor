import WebViewer from '@pdftron/webviewer';
import accountingData from '../accounting_data.json';

const app = document.getElementById('app');

WebViewer.Iframe({
  path: '/webviewer',
  // initialDoc: '/annual_financial_report.xlsx', // if we don't load this, an empty spreadsheet loads instead
  initialMode: WebViewer.Modes.SPREADSHEET_EDITOR,
  licenseKey: import.meta.env.VITE_PDFTRONKEY, // this comes from vite, should be your key (keep private)
}, app).then(instance => {

  const {
    documentViewer,
    SpreadsheetEditor: {
      SpreadsheetEditorManager,
      SpreadsheetEditorEditMode,
      Types }
  } = instance.Core;

  const spreadsheetEditorManager = documentViewer.getSpreadsheetEditorManager();
  const SpreadsheetEditorEvents = SpreadsheetEditorManager.Events;

  spreadsheetEditorManager.addEventListener(SpreadsheetEditorEvents.SPREADSHEET_EDITOR_READY, () => {

    spreadsheetEditorManager.setEditMode(
      SpreadsheetEditorEditMode.EDITING
    );

    const spreadsheetEditorDocument = documentViewer.getDocument().getSpreadsheetEditorDocument();

    const workbook = spreadsheetEditorDocument.getWorkbook();
    const font = workbook.createFont({
      fontFace: 'Times New Roman',
      pointSize: 12,
      color: 'black',
      bold: true,
    });

    const sheet = workbook.getSheetAt(0);
    const tableHeaders = Object.keys(accountingData[0]);
    const cellStyle = sheet.getRowAt(0).getCellAt(0).getStyle();
    cellStyle.font = font;
    cellStyle.horizontalAlignment = Types.HorizontalAlignment.CENTER;
    cellStyle.verticalAlignment = Types.VerticalAlignment.CENTER;
    cellStyle.backgroundColor = 'lightgreen';
    cellStyle.wrapText = Types.TextWrap.OVERFLOW;
    sheet.setColumnWidthInPixel(0, 200);

    for (let i = 0; i <= accountingData.length - 1; i++) {
      for (let j = 0; j <= tableHeaders.length - 1; j++) {
        const key = tableHeaders[j];
        const cellValue = accountingData[i][key];
        const cell =
          sheet
            .getRowAt(i)
            .getCellAt(j);
        typeof cellValue === "string" ? cell.setStringValue(cellValue) : cell.setNumericValue(cellValue);
        if (typeof cellValue === "string") {
          cell.setStyle(cellStyle);
        }
      }
    }
  });
});
