/** @OnlyCurrentDoc */
namespace DataSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Daten");
  export function clear(): void {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
    }
  }
  /// Add additional entries, keeps the worksheet sorted.
  export function add(entries: Entry.IEntry[]): void {
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
        { column: 4, ascending: false },
        // employee
        { column: 2, ascending: true },
       ]);
    }
  }
  /// Remove all entries with matching values
  export function removeMatching(date: Date, locationName: string, shiftName: string): void {
    // Given the sorting (see append), all relevant rows will be consecutive
    const allData = sheet.getDataRange();
    const withoutHeader = allData.offset(1, 0, allData.getNumRows() - 1);
    const data = withoutHeader.getValues();
    function matches(row: any[]) {
      return (DateUtils.equal(Values.asDate(row[0]), date) &&
        row[2] === locationName && row[3] === shiftName);
    }
    let firstRow = 0;
    while (firstRow < data.length && !matches(data[firstRow])) {
      firstRow++;
    }
    if (firstRow >= data.length) {
      // No matching entries.
      return;
    }
    Logger.log("firstRow: %s", firstRow);
    // At least one matching entry, find more...
    let rows = 1;
    while (firstRow + rows < data.length && matches(data[firstRow + rows])) {
      rows++;
    }
    // firstRow is 0 based, Add one because deleteRows is 1 based and one more
    // for the header
    sheet.deleteRows(firstRow + 2, rows);
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
    add(existingOutsideRange.concat(entries));
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
