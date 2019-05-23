/** @OnlyCurrentDoc */
namespace DataSheet {
  /// The data sheet stores entries in flattened fashion (aka one row per employee).
  interface Line extends Entry.Slot {
    employee: string;
  }

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Daten");

  export function initialSetup() {
    sheet.getRange("E2:E").setNumberFormat("[hh]:mm");
    sheet.getRange("F2:F").setNumberFormat("[hh]:mm");
    sheet.getRange("G2:G").setNumberFormat("[hh]:mm");
  }

  function entryToLines(e: Entry.IEntry): Line[] {
    return e.employees.map((employee) => ({
      date: e.date,
      location: e.location,
      shift: e.shift,
      employee,
    }));
  }

  export function clear(): void {
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clear();
    }
  }

  function addLines(lines: Line[]): void {
    const values = lines.map((l: Line) => {
      const wallclocktime = Interval.diff(l.shift.stop, l.shift.start);
      const worktime = Interval.diff(wallclocktime, l.shift.breakLength);

      return [l.date, l.employee, l.location.name, l.shift.name,
      l.shift.start.toHHMM(), l.shift.stop.toHHMM(), l.shift.breakLength.toHHMM(),
      worktime.getTotalMinutes(),
      ];
    });
    if (values.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, values.length, values[0].length).setValues(values);

      // Make sure the whole sheet is still correctly sorted
      const allData = sheet.getDataRange();
      const withoutHeader = allData.offset(1, 0, allData.getNumRows() - 1);
      withoutHeader.sort([
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

  /** Add additional entries, keeps the worksheet sorted.  Any existing entries on the same
   * slot are merged as well as any in the input.
   */
  export function add(entries: Entry.IEntry[]): void {
    addLines(Prelude.flattenArray(entries.map(entryToLines)));
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
    // At least one matching entry, find more...
    let rows = 1;
    while (firstRow + rows < data.length && matches(data[firstRow + rows])) {
      rows++;
    }
    // firstRow is 0 based, Add one because deleteRows is 1 based and one more
    // for the header
    sheet.deleteRows(firstRow + 2, rows);
  }

  function forEachLine(f: (e: Line) => void): void {
    const data = sheet.getDataRange().getValues();
    data.shift();
    data.forEach((row) => {
      const date = Values.asDate(row[0]);
      const employee = Values.asString(row[1]);
      const locationName = Values.asString(row[2]);
      const shiftName = Values.asString(row[3]);
      const shiftStart = Values.asInterval(row[4]);
      const shiftStop = Values.asInterval(row[5]);
      const shiftBreakLength = Values.asInterval(row[6]);
      const location = Prelude.unwrap(Locations.byName(locationName));
      const shift = Prelude.unwrap(Shifts.byName(shiftName));
      // TODO: validate shift settings?  Think about the potential mismatch
      f({ date, employee, location, shift });
    });
  }

  export function replaceRange(fromDate: Date, toDate: Date, entries: Entry.IEntry[]): void {
    const existingOutsideRange: Line[] = [];
    forEachLine((l) => {
      if (!(DateUtils.inRangeInclusive(l.date, fromDate, toDate))) {
        existingOutsideRange.push(l);
      }
    });
    clear();
    addLines(existingOutsideRange.concat(...entries.map(entryToLines)));
  }

  // Call f with non empty lists of all entries at the same date, location and shift
  export function forEach(f: (entry: Entry.IEntry) => void): void {
    let cur: Entry.IEntry | undefined;
    forEachLine((l: Line) => {
      if (cur === undefined) {
        cur = { date: l.date, location: l.location, shift: l.shift, employees: [l.employee] };
      } else if (Entry.sameSlot(cur, l)) {
        cur.employees.push(l.employee);
      } else {
        f(cur);
        cur = { date: l.date, location: l.location, shift: l.shift, employees: [l.employee] };
      }
    });
    if (cur !== undefined) {
      f(cur);
    }
  }
}
