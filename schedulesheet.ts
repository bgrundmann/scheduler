/** @OnlyCurrentDoc */
namespace ScheduleSheet {
  const INDEX_COLUMN = 5;
  const FIRST_ENTRY_COLUMN = 7;
  const FIRST_ENTRY_ROW = 3;
  const ROWS_PER_ENTRY = 2;
  const COLUMNS_PER_ENTRY = 2;

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Schedule");

  let dateRangeCache: { from: Date, until: Date } = null;

  // Return the date range as given by the 'von' -> 'bis' cells
  export function dateRange(): { from: Date, until: Date } {
    if (!dateRangeCache) {
      const dates = sheet.getRange(1, 2, 1, 3).getValues();
      const from = Values.get(dates, 0, 0, Values.asDate);
      const until = Values.get(dates, 0, 2, Values.asDate);
      dateRangeCache = { from, until };
    }
    return dateRangeCache;
  }

  function splitNames(entry: string): string[] {
    return entry.split(",").map((e) => e.trim()).filter((e) => e !== "");
  }
  // setup the part of the sheet on the left that lists employees and how much they work
  function setupEmployeeSection(): void {
    const employees = EmployeeSheet.all().map((e) => [ e.employee ]);
    sheet.getRange(FIRST_ENTRY_ROW - 1, 1, 1, 3).setValues([["Mitarbeiter", "Stunden", ""]]).setFontWeight("bold");
    sheet.getRange(FIRST_ENTRY_ROW, 1, employees.length, 1).setValues(employees);
    const oneCell =
      sheet.getRange(FIRST_ENTRY_ROW, 2)
      .setFormula('=SUMIFS(Daten!H$2:H; Daten!B$2:B; "="&A3; Daten!A$2:A; ">="&$B$1; Daten!A$2:A; "<="&$D$1)')
    .copyTo(sheet.getRange(FIRST_ENTRY_ROW, 2, employees.length, 1));
    const locs = Locations.all().map((c) => c.name);
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(locs).build();
    sheet.getRange(FIRST_ENTRY_ROW, 3, employees.length, 1).setDataValidation(rule);
    // temporarily set entries for sizing
    sheet.getRange(FIRST_ENTRY_ROW, 3, locs.length, 1).setValues(locs.map((l) => [l]));
    sheet.getRange(FIRST_ENTRY_ROW - 1, 1, employees.length + 1, 3).applyRowBanding();
    sheet.autoResizeColumns(1, 4);
    // now that sizing is done reset the values
    sheet.getRange(FIRST_ENTRY_ROW, 3, locs.length, 1).setValues(locs.map((l) => [""]));
  }
  // return the top left row the entry on the given date
  function entryRow(date: Date): number {
    return FIRST_ENTRY_ROW + DateUtils.daysBetween(dateRange().from, date) * ROWS_PER_ENTRY;
  }
  function entryColumn(date: Date, loc: Locations.ILocation): number {
    return FIRST_ENTRY_COLUMN + loc.ndx * (COLUMNS_PER_ENTRY + 1);
  }
  export function formulasEmployeeCount(date: Date, loc: Locations.ILocation): string[] {
    function countFormula(cell: string) {
      return "IF(ISBLANK(Schedule!" + cell + ");0;LEN(Schedule!" + cell +
        ")-LEN(SUBSTITUTE(Schedule!" + cell + ';",";""))+1)';
    }
    const row = entryRow(date);
    const col = entryColumn(date, loc);
    const wholeDay = countFormula(SheetUtils.a1(row, col));
    const firstHalf = countFormula(SheetUtils.a1(row + 1, col));
    const secondHalf = countFormula(SheetUtils.a1(row + 1, col + 1));
    return ["=" + wholeDay + "+" + firstHalf, "=" + wholeDay + "+" + secondHalf];
  }
  // Setup the sheet and copy the range of entries from the data sheet
  export function setup(fDate: Date, tDate: Date) {
    sheet.clear();
    sheet.setHiddenGridlines(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(FIRST_ENTRY_COLUMN - 2);
    dateRangeCache = null;
    sheet.getRange(1, 1, 1, 4).setValues([["Von", fDate, "bis", tDate]]);
    sheet.getRangeList(["A1", "C1"]).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRangeList(["B1", "D1"]).setNumberFormat("yyyy-mm-dd");
    setupEmployeeSection();
    // make sure columns on the right are sized properly
    Locations.all().forEach((loc, ndx) => {
      const col = FIRST_ENTRY_COLUMN + ndx * (COLUMNS_PER_ENTRY + 1);
      sheet.getRange(1, col).setValue(loc.name).setFontWeight("bold");
      sheet.setColumnWidth(col - 1, 10);
      sheet.setColumnWidth(col, 100);
      sheet.setColumnWidth(col + 1, 100);
    });

    // write the column of dates, draw the boxes and setup the formula for #people per entry
    DateUtils.forEachDay(dateRange().from, dateRange().until, (date) => {
      const row = entryRow(date);
      sheet.getRange(row, INDEX_COLUMN).setValue(date).setNumberFormat('ddd", "mmmm" "d');
      sheet.getRange(row, INDEX_COLUMN, 2, 1)
        .mergeVertically()
        .setBorder(true, true, true, true, false, false,  "#000000", SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment("middle");
      Locations.all().forEach((loc) => {
        const col = entryColumn(date, loc);

        sheet.getRange(row, col, 1, 2).mergeAcross();
        sheet.getRange(row, col, 2, 2)
          .setBorder(true, true, true, true, null, null, "#000000", SpreadsheetApp.BorderStyle.SOLID)
          .setBorder(null, null, null, null, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
      });
      if (DateUtils.isWeekend(date)) {
        sheet.getRange(row, INDEX_COLUMN, ROWS_PER_ENTRY, 1 + Locations.all().length * (COLUMNS_PER_ENTRY + 1))
          .setBackground(Config.WEEKEND_COLOR);
      }
    });
    sheet.autoResizeColumn(INDEX_COLUMN);
    const lastRow = entryRow(dateRange().until) + 2;
    sheet.getRange("B1").activate();

    const entryRange = sheet.getRange(FIRST_ENTRY_ROW, FIRST_ENTRY_COLUMN,
      lastRow - FIRST_ENTRY_ROW, Locations.all().length * (COLUMNS_PER_ENTRY + 1));
    const data = entryRange.getValues();

    // place entries from data sheet into schedule sheet
    DataSheet.forEachEntry((e: Entry.IEntry) => {
      if (DateUtils.inRangeInclusive(e.date, dateRange().from, dateRange().until)) {
        let row = entryRow(e.date) - FIRST_ENTRY_ROW;
        let col = entryColumn(e.date, e.location) - FIRST_ENTRY_COLUMN;
        const offset = e.shift.entryDisplayOffset;
        row += offset[0];
        col += offset[1];
        const existing = data[row][col];
        const newValue = existing === "" ? e.employee : existing + ", " + e.employee;
        data[row][col] = newValue;
      }
    });

    entryRange.setValues(data);
  }

  // Call f for each entry on the schedule sheet
  export function forEachEntry(f: (schedule: Entry.IEntry) => void): void {
    const entryRows = DateUtils.daysBetween(dateRange().from, dateRange().until) + 1;
    const dataRange = sheet.getRange(FIRST_ENTRY_ROW, FIRST_ENTRY_COLUMN,
      entryRows * ROWS_PER_ENTRY,
      Locations.all().length * (COLUMNS_PER_ENTRY + 1)).getValues();
    Locations.all().forEach ((loc) => {
      DateUtils.forEachDay(dateRange().from, dateRange().until, (date) => {
        const row = entryRow(date) - FIRST_ENTRY_ROW;
        const col = entryColumn(date, loc) - FIRST_ENTRY_COLUMN;
        const whole = splitNames(Values.get(dataRange, row, col, Values.asString));
        const firstHalf = splitNames(Values.get(dataRange, row + 1, col, Values.asString));
        const secondHalf = splitNames(Values.get(dataRange, row + 1, col + 1, Values.asString));
        const all =
            [ { shift : Shifts.whole, names : whole },
             { shift : Shifts.firstHalf, names : firstHalf },
             { shift : Shifts.secondHalf, names : secondHalf },
            ] ;
        all.forEach((a) => {
          a.names.forEach((employee) => {
            const entry = {
              date, employee, location : Locations.all()[loc.ndx], shift : a.shift,
            };
            f (entry);
          });
        });
      });
    });
  }
  // Get employees and locations to schedule as dictionary (from the left pane)
  export function employeesAndLocations() {
    const employeeCount = EmployeeSheet.all().length;
    const data = sheet.getRange(FIRST_ENTRY_ROW, 1, employeeCount, 3)
      .getValues()
      .filter((row) => row[2] !== "")
      .map((row) => ({
        employee : Values.asString(row[0]),
        location : Locations.byName(Values.asString(row[2])),
      }));
    return Prelude.makeDictionary(data, (d) => d.employee);
  }
}

function testThis() {
  Logger.clear();
  ScheduleSheet.setup(new Date("2019-04-11"), new Date("2019-05-22"));
}
