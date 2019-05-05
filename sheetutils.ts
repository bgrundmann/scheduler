/** @OnlyCurrentDoc */
namespace SheetUtils {
  export function createOrClearSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
    const ss = SpreadsheetApp.getActive(); const sheet = ss.getSheetByName(name);
    if (sheet === null) {
      return ss.insertSheet(name);
    } else {
      sheet.clear();
    }
    return sheet;
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
}
