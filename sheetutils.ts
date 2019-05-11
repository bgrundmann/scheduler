/** @OnlyCurrentDoc */
namespace SheetUtils {
  /** A decoded version of the onEdit event:
   * OnEditInsert means the cell was previously empty
   * OnEditChange means the cell was not empty
   * OnEditDelete means the cell is now empty
   */
  export interface OnEditInsert {
    kind: "insert";
    value: any;
  }

  export interface OnEditChange {
    kind: "change";
    value: any;
    oldValue: any;
  }

  export interface OnEditClear {
    kind: "clear";
    oldValue: any;
  }

  export interface OnEditMassChange {
    kind: "mass-change";
  }

  export type OnEditEvent = OnEditInsert | OnEditChange | OnEditClear | OnEditMassChange;

  /** Turn a google sheet onEdit Event into a typed event. */
  export function onEditEvent(event: GoogleAppsScript.Events.SheetsOnEdit): OnEditEvent {
    if (event.oldValue === undefined && event.value === undefined) {
      return { kind: "mass-change" };
    } else if (event.oldValue === undefined && event.value !== undefined) {
      return { kind: "insert", value: event.value };
    } else if (event.oldValue !== undefined && event.value.oldValue !== undefined) {
      return { kind: "clear", oldValue: event.oldValue };
    } else {
      return { kind: "change", oldValue: event.oldValue, value: event.value };
    }
  }

  export function createOrClearSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
    const ss = SpreadsheetApp.getActive(); const sheet = ss.getSheetByName(name);
    if (sheet === null) {
      return ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    return sheet;
  }
  /** Same as getDataRange but excludes the header.  Returns undefined if the resulting
   * range would be empty.
   */
  export function getDataRangeWithoutHeader(sheet: GoogleAppsScript.Spreadsheet.Sheet):
    GoogleAppsScript.Spreadsheet.Range | undefined {
    const allData = sheet.getDataRange();
    if (allData.getNumRows() > 1) {
      return allData.offset(1, 0, allData.getNumRows() - 1, allData.getNumColumns());
    } else {
      return undefined;
    }
  }
  export function deleteSheetByNameIfExists(name: string): void {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(name);
    if (sheet !== null) {
      ss.deleteSheet(sheet);
    }
  }
  export function convertColumnLetterToNumber(a: string): number {
    let res = 0;
    for (let i = 0; i < a.length; i++) {
      res = res * 26 + (a.charCodeAt(i) + 1 - "A".charCodeAt(0));
    }
    return res;
  }
  export function convertNumberToColumnLetter(n: number): string {
    const res = [];
    while (n > 0) {
      const digit = n % 26;
      if (digit === 0) {
        res.push("Z");
        n = Math.floor(n / 26 - 1);
      } else {
        res.push(String.fromCharCode("A".charCodeAt(0) + digit - 1));
        n = Math.floor(n / 26);
      }
    }
    res.reverse();
    return res.join("");
  }
  export function a1(row: number, col: number): string {
    return convertNumberToColumnLetter(col) + String(row);
  }

  export function buildRichTexts(runs: Array< { text: string, style: GoogleAppsScript.Spreadsheet.TextStyle } >):
  GoogleAppsScript.Spreadsheet.RichTextValue {
    const b = SpreadsheetApp.newRichTextValue();
    const completeText = runs.map(({ text, style }) => text).join("");
    b.setText(completeText);
    let off = 0;
    runs.forEach( ({ text, style }) => {
      b.setTextStyle(off, off + text.length, style);
      off += text.length;
    });
    return b.build();
  }

  /// is the given range a single cell?
  export function isCell(r: GoogleAppsScript.Spreadsheet.Range): boolean {
    return r.getNumColumns() === 1 && r.getNumRows() === 1;
  }

  /** Prompt user for a date.  Returns undefined if user pressed cancel. */
  export function askForDate(prompt: string): Date|undefined {
    const ui = SpreadsheetApp.getUi();
    while (true) {
      const result = ui.prompt(prompt, ui.ButtonSet.OK_CANCEL);
      const button = result.getSelectedButton();
      if (button === ui.Button.OK) {
        const res = DateUtils.parseISODate(result.getResponseText());
        if (res) {
          return res;
        } else {
          ui.alert("Habe deine Eingabe nicht verstanden.  Bitte YYYY-MM-DD (zBsp 1981-07-14)");
        }
      } else if (button === ui.Button.CANCEL || button === ui.Button.CLOSE) {
        return undefined;
      }
    }
  }

  /** Create a copy of the sheet named sourceName called backupName.  If a sheet of such
   * a name already exists it's content is replaced.
   */
  export function backupSheet(sourceName: string, backupName: string): void {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SheetUtils.deleteSheetByNameIfExists(backupName);
    const sheet = ss.getSheetByName(sourceName).copyTo(ss);
    sheet.setName(backupName);
  }

  /** Restore a backup of a sheet as created by backupSheet, to the sheet called
   * dstName.  Returns false if there is no backup of the given name.
   */
  export function restoreSheet(backupName: string, dstName: string): boolean {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupSheet = ss.getSheetByName(backupName);
    if (!backupSheet) { return false; }
    const dst = ss.getSheetByName(dstName);
    dst.clear();
    backupSheet.getDataRange().copyTo(dst.getRange(1, 1));
    return true;
  }

  export function autoResizeColumns(sheet: GoogleAppsScript.Spreadsheet.Sheet,
                                    startColumn: number, numColumns: number, minWidthPixels?: number) {
    sheet.autoResizeColumns(startColumn, numColumns);
    if (minWidthPixels) {
      for (let c = startColumn; c < startColumn + numColumns; c++) {
        if (sheet.getColumnWidth(c) < minWidthPixels) {
          sheet.setColumnWidth(c, minWidthPixels);
        }
      }
    }
  }
}
