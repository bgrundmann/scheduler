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

      // Make sure the whole sheet is still correctly sorted
      const allData = sheet.getDataRange();
      const withoutHeader = allData.offset(1, 0, allData.getNumRows() - 1);
      withoutHeader.sort( [
        // date
        { column: 1, ascending: true },
        // location
        { column: 3, ascending: true },
        // shift
        { column: 4, ascending: true },
        // employee
        { column: 2, ascending: true },
       ]);
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
      const location = Prelude.unwrap(Locations.byName(locationName));
      const shift = Prelude.unwrap(Shifts.byName(shiftName));
      // TODO: validate shift settings?  Think about the potential mismatch
      f({ date, employee, location, shift });
    });
  }

  // Call f with non empty lists of all entries at the same date, location and shift
  export function forEachEntryGrouped(f: (entries: Entry.IEntry[]) => void): void {
    let cur: Entry.IEntry[] = [];
    forEachEntry((e: Entry.IEntry) => {
      if (cur.length === 0 ||
        (cur[0].date.getTime() === e.date.getTime() && 
          cur[0].location.name === e.location.name &&
          cur[0].shift.name === e.shift.name)) {
        cur.push(e);
      } else {
        f(cur);
        cur = [e];
      }
    });
    if (cur.length !== 0) {
      f(cur);
    }
  }
}
