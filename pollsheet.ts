namespace PollSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("UmfrageAlsTabelle");

  interface IPoll {
    employee: string;
    date: Date;
    shift: Shifts.Shift;
  }

  export function forEach(f: (poll: IPoll) => void): void {
    const dataRange = sheet.getDataRange().getValues();
    const rows = dataRange.length;
    for (let row = 1; row < rows; row++) {
      const start = Values.get(dataRange, row, 2, Values.asInterval);
      const end = Values.get(dataRange, row, 3, Values.asInterval);
      let inferredBreakLength = Interval.zero;
      if (Interval.diff(end, start).getHours() >= 7) {
        inferredBreakLength = Interval.hhmm(1, 0);
      }
      const shift = Shifts.create(start, end, inferredBreakLength);
      f({
        employee: Values.get(dataRange, row, 0, Values.asString),
        date: Values.get(dataRange, row, 1, Values.asDate), shift,
      });
    }
  }

  // Each employee only once per date and with the longest available shift
  export function forEachUnique(f: (poll: IPoll) => void): void {
    let last: IPoll | undefined;
    forEach((p) => {
      if (!last || last.employee !== p.employee || !DateUtils.equal(last.date, p.date)) {
        last = p;
        f(p);
      }
    });
  }
}
