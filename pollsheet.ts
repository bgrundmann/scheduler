namespace PollSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("UmfrageAlsTabelle");

  interface IPoll {
    employee: string;
    date: Date;
    start: Date;
    end: Date;
    shift: Shifts.IShift;
  }

  export function forEach(f: (poll: IPoll) => void): void {
    const dataRange = sheet.getDataRange().getValues();
    const rows = dataRange.length;
    for (let row = 1; row < rows; row++) {
      const shift = Prelude.unwrap(Shifts.byName(Values.get(dataRange, row, 4, Values.asString)));
      f({ employee: Values.get(dataRange, row, 0, Values.asString),
        date: Values.get(dataRange, row, 1, Values.asDate),
        start : Values.get(dataRange, row, 2, Values.asDate),
        end : Values.get(dataRange, row, 3, Values.asDate),
        shift });
    }
  }

  // Each employee only once per date and with the longest available shift
  export function forEachUnique(f: (poll: IPoll) => void): void {
    function len(x: IPoll) { return x.shift.stop - x.shift.start; }
    let last: IPoll|undefined;
    forEach((p) => {
      if (last === undefined) {
        last = p;
      } else if (last.employee === p.employee && last.date.getTime() === p.date.getTime()) {
        // choose the longer one
        if (len(p) > len(last)) {
          last = p;
        }
      } else {
        f(last);
        last = p;
      }
    });
    if (last) {
      f(last);
    }
  }
}
