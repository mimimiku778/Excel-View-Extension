import OfficeDocumentSettings from "./OfficeDocumentSettings";

const PREVIOUS_KEY = "highlightPrevious";
const COLOR_KEY = "highlightColor";

namespace HighlightRowStorage {
  export type Previous = {
    [key: string]: { address: string; id: number; errorCount: number }[];
  };

  // Store the address and id of the conditional format in local storage
  export async function setPrevious(worksheetName: string, address: string, id: number) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
    if (!previous[worksheetName]) previous[worksheetName] = [];

    previous[worksheetName].push({ address, id, errorCount: 0 });

    await OfficeDocumentSettings.set(PREVIOUS_KEY, JSON.stringify(previous));
  }

  // Retrieve the address and id of the conditional format from local storage
  export function getPreviousAll(): Previous {
    return JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
  }

  // Retrieve the address and id of the conditional format from local storage
  export async function deletePrevious(worksheetName: string, adress: string) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
    if (!previous[worksheetName]) return;

    previous[worksheetName] = previous[worksheetName].filter((item) => item.address !== adress);

    await OfficeDocumentSettings.set(PREVIOUS_KEY, JSON.stringify(previous));
  }

  /**
   * Increment the error count in the local storage.
   * 
   * @returns The number of errors.
   */
  export async function incrementErrorCount(worksheetName: string, adress: string) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
    if (!previous[worksheetName]) return 0;

    let errorCount = 0;
    previous[worksheetName] = previous[worksheetName].map((item) => {
      if (item.address === adress) {
        errorCount = item.errorCount + 1;
        return { ...item, errorCount };
      }
      return item;
    });

    await OfficeDocumentSettings.set(PREVIOUS_KEY, JSON.stringify(previous));
    return errorCount;
  }

  /**
   * Check if the row exists in the local storage.
   * 
   * @returns The id of the conditional format if the row exists, otherwise false. 
   */
  export function rowExists(worksheetName: string, adress: string) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
    if (!previous[worksheetName]) return false;

    for (const item of previous[worksheetName]) {
      if (item.address === adress) {
        return item.id;
      }
    }

    return false;
  }

  // Retrieve the address and id of the conditional format from local storage
  export async function deleteRow(worksheetName: string, adress: string, id: number) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};
    if (!previous[worksheetName]) throw new Error("The worksheet does not exist.");

    let deleted = false;
    previous[worksheetName] = previous[worksheetName].filter((item) => {
      if (item.address === adress && item.id === id) {
        deleted = true;
        return false;
      } else {
        return true;
      }
    });

    if (!deleted) throw new Error("The row does not exist.");

    await OfficeDocumentSettings.set(PREVIOUS_KEY, JSON.stringify(previous));
  }

  // Retrieve the address and id of the conditional format from local storage
  export async function renameWorksheet(nameAfter: string, nameBefore: string) {
    const previous: Previous = JSON.parse(OfficeDocumentSettings.get(PREVIOUS_KEY)) ?? {};

    previous[nameAfter] = previous[nameBefore] ?? [];
    delete previous[nameBefore];

    await OfficeDocumentSettings.set(PREVIOUS_KEY, JSON.stringify(previous));
  }

  // Retrieve the address and id of the conditional format from local storage
  export function xColor(color?: string): string | null {
    if (color !== undefined) {
      localStorage.setItem(COLOR_KEY, color);
      return color;
    }

    return localStorage.getItem(COLOR_KEY);
  }
}

export default HighlightRowStorage;