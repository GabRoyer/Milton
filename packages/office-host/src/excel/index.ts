async function ensureExcelReady(): Promise<void> {
  if (typeof Office === "undefined") {
    throw new Error("Office.js did not load.");
  }

  await Office.onReady();

  if (typeof Excel === "undefined") {
    throw new Error("Excel APIs are not available in this host.");
  }
}

export async function insertText(text: string): Promise<void> {
  await ensureExcelReady();

  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");

    range.values = [[text]];
    range.format.autofitColumns();

    await context.sync();
  });
}
