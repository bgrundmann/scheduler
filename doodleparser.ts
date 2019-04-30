/** @OnlyCurrentDoc */
namespace DoodleParser {
  // TODO: parse the year from the sheet
  const year = 2019;

  const germanMonthNames = (() => {
    const res: Record<string, number> = {};
    ["Januar", "Februar", "März", "April", "Mai",
      "Juni", "Juli", "August", "September",
      "Oktober", "November", "Dezember"].forEach((name, ndx) => {
      res[name] = ndx;
    });
    return res;
  })();

  interface ICell {
    column: number;
    lastColumn: number;
    value: any;
  }

  // Given a row of potentially merged cells, return an array of column and value objects
  // (column being the column of the first cell and value the value in that cell)
  function parseMergedRow(sheet: GoogleAppsScript.Spreadsheet.Sheet,
                          row: number, column: number, lastColumn: number): ICell[] {
    const result = [];
    const ranges = sheet.getRange(row, column, 1, lastColumn - column).getMergedRanges();
    ranges.sort((a, b) => a.getColumn() - b.getColumn() );
    let last = column - 1;
    if (ranges.length === 0) {
      for (let c = last + 1; c <= lastColumn; c++) {
        result.push( { column : c, lastColumn : c, value : sheet.getRange(row, c).getValue() } );
      }
    }
    // TODO: handle single cells at beginning or end
    ranges.forEach((r) => {
      // the merged cells aren't in the array returned by getMergedRanges, so
      // we need to get them in another way.
      for (let c = last + 1; c < r.getColumn(); c++) {
        result.push( { column : c, lastColumn : c, value : sheet.getRange(row, c).getValue() } );
      }
      result.push( { column : r.getColumn(), lastColumn : r.getLastColumn(), value : r.getValue() } );
      last = r.getLastColumn();
    });
    return result;
  }

  function betweenInclusive(x: number, low: number, high: number) {
    return low <= x && x <= high;
  }

  const monthAndYearRegex = /^([^ ]+) ([0-9]+)$/;
  const timeRegex = /^([0-9]+):([0-9]+) – ([0-9]+):([0-9]+)$/;
  const dayRegex = /^([^ ]+) ([0-9]+)$/;

  interface IShift {
    year: number;
    month: number;
    day: number;
    timeStartHour: number;
    timeStartMinute: number;
    timeEndHour: number;
    timeEndMinute: number;
  }

  function forEachShift(monthAndYears: any[], days: any[], times: any[],
                        f: (shift: IShift, column: number) => void): void {
    let month = 0;
    let day = 0;
    let time = 0;
    const lastColumn = monthAndYears[monthAndYears.length - 1].lastColumn;

    for (let column = monthAndYears[0].column; column <= lastColumn; column++) {
      Logger.log("[1] column = %s, month = %s (%s), day = %s (%s), time = %s (%s)",
        column, month, monthAndYears[month], day, days[day], time, times[time]);
      if (!betweenInclusive(column, monthAndYears[month].column, monthAndYears[month].lastColumn)) {
        month++;
      }
      if (!betweenInclusive(column, days[day].column, days[day].lastColumn)) {
        day++;
      }
      if (!betweenInclusive(column, times[time].column, times[time].lastColumn)) {
        time++;
      }
      Logger.log("[2] column = %s, month = %s (%s), day = %s (%s), time = %s (%s)",
        column, month, monthAndYears[month], day, days[day], time, times[time]);
      let r = monthAndYearRegex.exec(monthAndYears[month].value);
      if (r == null) {
        // problem
      }
      // Parser would be more robust if this lookup failed with an error message as well
      const monthValue = germanMonthNames[r[1]];
      const year2 = Number(r[2]);

      // Parser would be more robust if this lookup failed with an error message as well
      r = dayRegex.exec(days[day].value);
      const dayValue = Number(r[2]);

      // Parser would be more robust if this lookup failed with an error message as well
      r = timeRegex.exec(times[time].value);
      const timeStartHour = Number(r[1]);
      const timeStartMinute = Number(r[2]);
      const timeEndHour = Number(r[3]);
      const timeEndMinute = Number(r[4]);
      f( { year : year2, month : monthValue, day : dayValue,
        timeStartHour , timeStartMinute ,
        timeEndHour , timeEndMinute } , column);
    }
  }

  export function parse() {
    Logger.clear();
    const ss = SpreadsheetApp.getActive();
    const doodle = ss.getSheetByName("Umfrage");
    const employeeDict = EmployeeSheet.byAliasAndHandle();
    Logger.log(employeeDict);
    const monthAndYears = parseMergedRow(doodle, 4, 2, doodle.getLastColumn());
    const days = parseMergedRow(doodle, 5, 2, doodle.getLastColumn());
    const times = parseMergedRow(doodle, 6, 2, doodle.getLastColumn());

    const values = doodle.getRange(7, 1, doodle.getLastRow() - 7 - 3, doodle.getLastColumn() - 1).getValues();
    const result: any[][] = [];
    forEachShift(monthAndYears, days, times, (d, c) => {
      // NOTE: forEachShift column is absolute column on sheet (with 1 being A)
      // But values array is 0 based
      for (const row of values) {
          const parsedName = String(row[0]);
          // FIXME: Error handling here
          const name = employeeDict[parsedName].employee;
          const ok = row[c - 1] === "OK";
          if (ok) {
            const date  = new Date(d.year, d.month, d.day);
            const start = new Date(d.year, d.month, d.day, d.timeStartHour, d.timeStartMinute);
            const end = new Date(d.year, d.month, d.day, d.timeEndHour, d.timeEndMinute);
            let shift = null;
            // FIXME: Replace by something that takes actual times as defined by shifts into account
            if ((d.timeStartHour === 9 || d.timeStartHour === 10) && d.timeEndHour === 14) {
              shift = Shifts.byName("Vormittags");
            } else if ((d.timeStartHour === 9 || d.timeStartHour === 10) && d.timeEndHour > 14) {
              shift = Shifts.byName("Ganztags");
            } else if (d.timeStartHour === 13) {
              shift = Shifts.byName("Nachmittags");
            }
            result.push([name, date, start, end, shift.name]);
          }
      }
    });
    const data = SheetUtils.createOrClearSheetByName("UmfrageAlsTabelle");
    data.clear();
    data.appendRow(["Mitarbeiter", "Tag", "Anfang", "Ende", "Schicht"]);
    data.getRange(1, 1, 1, 5).setFontWeight("bold");
    if (result.length > 0) {
      data.getRange(2, 1, result.length, result[0].length)
      .setValues(result)
      .offset(1, 0, result.length - 1).sort([{column: 2, ascending: true}, {column: 1, ascending: true}]);
    }
    data.getRange(2, 3, result.length, 2).setNumberFormat("hh:mm");
  }
}

function test() {
  DoodleParser.parse();
}
