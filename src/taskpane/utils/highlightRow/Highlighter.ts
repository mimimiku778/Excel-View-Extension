import Storage from "../../../storage/HighlightRowStorage";
import Cleaner from "./HighlightCleaner";
import { SetError } from "../../hooks/useErrorSnackbar";

const CONDITIONAL_FORMAT_FORMULA = '=CELL("ROW")';
const SELECT_EVENT_TIMEOUT_MS = 30;
const CHANGE_COLOR_TIMEOUT_MS = 100;

export type HighlighterOption = { keepSelection: boolean; xColor: string };
type TimeoutId = { id: NodeJS.Timeout | null | number };

export default class Highlighter {
  private static selectEventTimeoutId: TimeoutId = { id: null };
  private static changeColorTimeoutId: TimeoutId = { id: null };

  public constructor(
    private option: HighlighterOption,
    private setError: SetError,
  ) {
  }

  public static create(setError: SetError, option: HighlighterOption) {
    return new Highlighter(option, setError);
  }

  public async updateRow() {
    await this.callHighlighter(Highlighter.selectEventTimeoutId, SELECT_EVENT_TIMEOUT_MS);
  }

  public async changeColor() {
    await this.callHighlighter(Highlighter.changeColorTimeoutId, CHANGE_COLOR_TIMEOUT_MS);
  }

  // Set the conditional format to highlight the active row
  private async callHighlighter(timeoutId: TimeoutId, timeoutMs: number) {
    // Check if the timeout is already set
    if (!timeoutId.id) {
      timeoutId.id = 1;
      await this.executeHighlight();
      timeoutId.id = setTimeout(() => undefined, timeoutMs);
    } else {
      // Set a timeout to avoid multiple calls
      clearTimeout(timeoutId.id);
      timeoutId.id = setTimeout(this.executeHighlight.bind(this), timeoutMs);
    }
  }

  private async executeHighlight() {
    if (!this.option.keepSelection) {
      // Clear the previous conditional format
      await Cleaner.create(this.setError).clearPrevious();
    }

    await Excel.run(async (context) => {
      // Get the active worksheet and the selected range
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const rangeArea = context.workbook.getSelectedRanges();
      rangeArea.load("cellCount");
      sheet.load("name");
      await context.sync();

      // Check if the selected range is too large
      if (rangeArea.cellCount > 1638400 || rangeArea.cellCount < 1) {
        return;
      }

      // Get the selected rows
      rangeArea.areas.load("items");
      await context.sync();

      // Highlight the active row
      for (const item of rangeArea.areas.items) {
        if (this.option.keepSelection) {
          await this.processHighlightEachRow(context, item.getEntireRow(), sheet.name);
        } else {
          await this.processHighlight(context, item.getEntireRow(), sheet.name);
        }
      }
    }).catch((error) => {
      this.setError([error.message, error.code, "executeHighlight"]);
    });
  }

  private async processHighlight(context: Excel.RequestContext, range: Excel.Range, sheetName: string) {
    // Highlight the active row
    const conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.custom.rule.formula = CONDITIONAL_FORMAT_FORMULA;
    conditionalFormat.custom.format.fill.color = this.option.xColor;

    // Save the highlighted row to the local storage
    range.conditionalFormats.load("count");
    range.load("address");
    const count = range.conditionalFormats.getCount();

    try {
      await context.sync();

      await Storage.setPrevious(sheetName, range.address, count.value - 1);
    } catch (error) {
      this.setError([error.message, error.code, "processHighlight"]);
    }
  }

  private async processHighlightEachRow(context: Excel.RequestContext, range: Excel.Range, sheetName: string) {
    try {
      range.load("rowCount");
      await context.sync();

      // Get the rows      
      const rows: { row: Excel.Range, adress: string, id: number | false }[] = [];
      for (let i = 0; i < range.rowCount; i++) {
        const row = range.getRow(i);
        row.load("address");
        await context.sync();

        rows.push({ row, adress: row.address, id: Storage.rowExists(sheetName, row.address) });
      }

      const selectedHighlight = rows.filter((row) => row.id === false);
      if (selectedHighlight.length) {
        for (const row of selectedHighlight) {
          await this.processHighlight(context, row.row, sheetName);
        }
      } else {
        for (const row of rows) {
          await Cleaner.create(this.setError).clearRow(context, row.row, sheetName, row.adress, row.id as number);
        }
      }
    } catch (error) {
      this.setError([error.message, error.code, "processHighlightEachRow"]);
    }
  }
}