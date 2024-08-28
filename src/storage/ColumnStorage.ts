import OfficeDocumentSettings from "./OfficeDocumentSettings";

const COLUMN_KEY = "column";

namespace ColumnStorage {
  export type Profiles = {
    [key: string]: number[][];
  };
  
  // Add the column profile to the column profile storage
  export async function addProfile(worksheetName: string, columns: number[]): Promise<number> {
    const store: Profiles = JSON.parse(OfficeDocumentSettings.get(COLUMN_KEY)) ?? {};
    if (!store[worksheetName]) store[worksheetName] = [];

    const index = store[worksheetName].push(columns);
    await OfficeDocumentSettings.set(COLUMN_KEY, JSON.stringify(store));
    
    return index;
  }

  // Get the column profile from the column profile storage
  export function getProfile(worksheetName: string): Profiles[keyof Profiles] {
    const store: Profiles = JSON.parse(OfficeDocumentSettings.get(COLUMN_KEY)) ?? {};
    
    return store[worksheetName] ?? [];
  }

  // Rename the worksheet in the column profile storage
  export async function renameWorksheet(nameAfter: string, nameBefore: string) {
    const store: Profiles = JSON.parse(OfficeDocumentSettings.get(COLUMN_KEY)) ?? {};

    store[nameAfter] = store[nameBefore] ?? [];
    delete store[nameBefore];

    await OfficeDocumentSettings.set(COLUMN_KEY, JSON.stringify(store));
  }

  // Delete the column profile from the column profile storage
  export async function deleteProfile(worksheetName: string, index: number) {
    const store: Profiles = JSON.parse(OfficeDocumentSettings.get(COLUMN_KEY)) ?? {};
    if (!store[worksheetName]) return;

    store[worksheetName] = store[worksheetName].filter((_, i) => i !== index);
    await OfficeDocumentSettings.set(COLUMN_KEY, JSON.stringify(store));
  }
}

export default ColumnStorage;

