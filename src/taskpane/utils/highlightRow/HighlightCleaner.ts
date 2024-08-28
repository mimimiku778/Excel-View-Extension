import Storage from "../../../storage/HighlightRowStorage";
import { SetError } from "../../hooks/useErrorSnackbar";

const MAX_ERROR_COUNT = 2;

/**
 * Clear the previously highlighted rows
 */
export default class HighlightCleaner {
  public constructor(
    private setError: SetError
  ) {
  }

  public static create(setError: SetError) {
    return new HighlightCleaner(setError);
  }

  // Clear the previous conditional format
  public async clearPrevious() {
    await Excel.run(async (context) => {
      // Get all the worksheets
      const worksheets = context.workbook.worksheets;
      worksheets.load("items");
      await context.sync();

      // Iterate through the worksheets
      const previous = Storage.getPreviousAll();
      for (const worksheetName of Object.keys(previous)) {
        // Get the worksheet
        const worksheet = worksheets.items.find((item) => item.name === worksheetName);
        if (!worksheet) continue;

        // Clear the previous conditional format
        await this.clearPreviousProcess(context, previous, worksheet);
      }
    }).catch((error) => {
      this.setError([error.message, error.code, "clearPrevious"]);
    });
  }

  private async clearPreviousProcess(
    context: Excel.RequestContext,
    previous: Storage.Previous,
    worksheet: Excel.Worksheet
  ) {
    // Remove the highlighted row from the worksheet
    for (const { address, id } of previous[worksheet.name].reverse()) {
      const conditionalFormat = worksheet.getRange(address).getEntireRow().conditionalFormats.getItemAt(id);
      conditionalFormat.delete();
    }

    try {
      // Save the changes
      await context.sync();

      // Remove the highlighted row from the local storage
      for (const { address } of previous[worksheet.name]) {
        await Storage.deletePrevious(worksheet.name, address);
      }
    } catch (error) {
      let deletedCount = 0;

      // Increment the error count in the local storage
      for (const { address } of previous[worksheet.name]) {
        const errorCount = await Storage.incrementErrorCount(worksheet.name, address);

        // If the error count exceeds the limit, remove the highlighted row from the local storage
        if (errorCount >= MAX_ERROR_COUNT) {
          await Storage.deletePrevious(worksheet.name, address);
          deletedCount++
        }
      }

      // Display an error message
      if (deletedCount) {
        this.setError([`The formula to highlight may have already been deleted. (${error.message})`, error.code]);
      }
    }
  }
}