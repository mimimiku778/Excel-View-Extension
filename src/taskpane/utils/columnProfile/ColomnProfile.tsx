import { SetError } from "../../hooks/useErrorSnackbar";

namespace ColumnProfile {
  // Get the active worksheet name
  export async function getActiveSheetName(
    setError: SetError,
    callback: (activeSheetName: string) => Promise<void> | void
  ) {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      await context.sync();

      await callback(sheet.name);
    }).catch((error) => {
      setError([error.message, error.code, "getActiveSheetName"]);
    });
  }

  /**
   * Get the column width of the range
   */
  export async function getColumnsFromWorksheet(
    rangeAddress: string,
    setError: SetError,
    callback: (activeSheetName: string, columns: number[]) => Promise<void> | void
  ) {
    await Excel.run(async (context) => {
      // Get the active worksheet
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(rangeAddress);
      range.load("columnCount");
      sheet.load("name");
      await context.sync();

      // Get the column width
      const columns: Excel.Range[] = [];
      for (let i = 0; i < range.columnCount; i++) {
        const column = range.getColumn(i);
        column.format.load("columnWidth");
        columns.push(column);
      }

      await context.sync();

      await callback(
        sheet.name,
        columns.map((column) => column.format.columnWidth)
      );
    }).catch((error) => {
      setError([error.message, error.code, "getColumnsFromWorksheet"]);
    });
  }

  /**
   * Set the column width of the range
   */
  export async function setColumnsToWorksheet(
    columns: number[],
    rangeAddress: string,
    setError: SetError,
    callback?: () => Promise<void> | void
  ) {
    await Excel.run(async (context) => {
      // Set the column width
      const range = context.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);
      columns.forEach((columnWidth, i) => {
        const column = range.getColumn(i);
        column.format.columnWidth = columnWidth;
      });

      await context.sync();

      callback && (await callback());
    }).catch((error) => {
      setError([error.message, error.code, "setColumnsToWorksheet"]);
    });
  }
}

export default ColumnProfile;
