namespace OfficeDocumentSettings {
  // Set the value in the localStorage
  export async function set(key: string, value: string) {
    await new Promise((resolve, reject) => {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve("Settings saved successfully");
        } else {
          reject(new Error("Unable to save settings to localStorage"));
        }
      });
    });
  }

  // Get the value from the localStorage
  export function get(key: string): string | null {
    return Office.context.document.settings.get(key);
  }
}

export default OfficeDocumentSettings;