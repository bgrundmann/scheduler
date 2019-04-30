/** @OnlyCurrentDoc */
namespace DataSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Daten");
  export function clear(): void {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
    }
  }
  export function append(entries: Entry.IEntry[]): void {
    const values = entries.map((e) => {
      return [ e.date, e.employee, e.location.name, e.shift.name,
        e.shift.start, e.shift.stop, e.shift.breakLength,
        e.shift.stop - e.shift.start - e.shift.breakLength ];
    });
    if (values.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);
    }
  }
  export function replaceRange(fromDate: Date, toDate: Date, entries: Entry.IEntry[]): void {
    const existingOutsideRange: Entry.IEntry[] = [];
    forEachEntry((e) => {
      if (!(DateUtils.inRangeInclusive(e.date, fromDate, toDate))) {
        Logger.log("%s <= %s <= %s", fromDate, e.date, toDate);
        existingOutsideRange.push(e);
      }
    });
    clear();
    append(existingOutsideRange.concat(entries));
  }
  export function forEachEntry(f: (e: Entry.IEntry) => void): void {
    const data = sheet.getDataRange().getValues();
    data.shift();
    data.forEach((row) => {
      const date = Values.asDate(row[0]);
      const employee = Values.asString(row[1]);
      const locationName = Values.asString(row[2]);
      const shiftName = Values.asString(row[3]);
      const shiftStart = Values.asNumber(row[4]);
      const shiftStop = Values.asNumber(row[5]);
      const shiftBreakLength = Values.asNumber(row[6]);
      Logger.log("DataSheet: %s", locationName);
      const location = Prelude.unwrap(Locations.byName(locationName));
      const shift = Prelude.unwrap(Shifts.byName(shiftName));
      // TODO: validate shift settings?  Think about the potential mismatch
      f({ date, employee, location, shift });
    });
  }
}
