/** @OnlyCurrentDoc */
namespace ScheduleSheet {
  const INDEX_COLUMN = 5;
  const FIRST_ENTRY_COLUMN = 7;
  const FIRST_ENTRY_ROW = 3;
  const ROWS_PER_ENTRY = 2;
  const COLUMNS_PER_ENTRY = 2;
  export const NAME = "Schedule";

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName(ScheduleSheet.NAME);

  let dateRangeCache: { from: Date, until: Date }|undefined;

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

  export function validateEntry(entry: string): boolean {
    const employees = EmployeeSheet.byAliasAndHandle();
    return entry.split(",").map((e) => e.trim()).every((e) => {
      const emp = employees[e];
      return emp && emp.employee === e;
    });
  }

  // return the top left row the entry on the given date
  function entryRow(date: Date): number {
    return FIRST_ENTRY_ROW + DateUtils.daysBetween(dateRange().from, date) * ROWS_PER_ENTRY;
  }

  function entryColumn(date: Date, loc: Locations.ILocation): number {
    return FIRST_ENTRY_COLUMN + loc.ndx * (COLUMNS_PER_ENTRY + 1);
  }

  // TODO: Rename cellToEntry
  function cellToEntry(row: number, column: number):
    { date: Date, location: Locations.ILocation, shift: Shifts.IShift } | undefined {
    if (row < FIRST_ENTRY_ROW) {
      return undefined;
    }
    if (column < FIRST_ENTRY_COLUMN) {
      return undefined;
    }
    const dr = dateRange();
    const date = DateUtils.addDays( dr.from, Math.floor((row - FIRST_ENTRY_ROW) / 2) );
    if (!DateUtils.inRangeInclusive(date, dr.from, dr.until)) {
      return undefined;
    }
    const locations = Locations.all();
    const locNdx = Math.floor((column - FIRST_ENTRY_COLUMN) / 3);
    const vpart = (row - FIRST_ENTRY_ROW) % 2;
    const hpart = (column - FIRST_ENTRY_COLUMN) % 3;
    if (!Prelude.inRangeInclusive(locNdx, 0, locations.length - 1)) {
      return undefined;
    }
    // are we on the empty columns between location columns?
    if (hpart === 2) {
      return undefined;
    }
    let shift: Shifts.IShift;
    if (hpart === 0 && vpart === 1) {
      shift = Shifts.firstHalf;
    } else if (hpart === 1 && vpart === 1) {
      shift = Shifts.secondHalf;
    } else if (vpart === 0) {
      shift = Shifts.whole;
    } else {
      // Do not know what is going on
      Logger.log("cellToEntry bug? (row=%s) (column=%s)", row, column);
      return undefined;
    }
    return { date, location: locations[locNdx], shift };
  }

  // get the range used to store all the data
  function getEntriesRange(): GoogleAppsScript.Spreadsheet.Range {
    const entryRows = DateUtils.daysBetween(dateRange().from, dateRange().until) + 1;
    return sheet.getRange(FIRST_ENTRY_ROW, FIRST_ENTRY_COLUMN,
      entryRows * ROWS_PER_ENTRY,
      Locations.all().length * (COLUMNS_PER_ENTRY + 1));
  }

  function forEachDayOnSheet(f: (date: Date) => void): void {
    DateUtils.forEachDay(dateRange().from, dateRange().until, f);
  }

  // function highlightEntries(): void {
  //   const range = getEntriesRange();
  //   const richTexts = range.getRichTextValues();
  //   forEachDayOnSheet((date) => {
  //     const row = entryRow(date);
  //     const rich = richTexts[row - FIRST_ENTRY_ROW][0];
  //     const e = rich.copy();
  //   });
  // }

  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  const normalStyle = SpreadsheetApp.newTextStyle().build();
  const errorStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").setBold(true).build();

  // TODO: Need better name and type for EntryGroup
  function layoutEntryGroup(entries: Entry.IEntry[]): GoogleAppsScript.Spreadsheet.RichTextValue {
    const elements =
      entries.map((e: Entry.IEntry) => {
        let style = normalStyle;
        switch (e.employee) {
          case undefined:
            style = normalStyle;
            break;

          case "not-in-poll":
            style = boldStyle;
            break;

          case "unknown-employee":
            style = errorStyle;
            break;
        }
        return { text: e.employee, style };
      }).intersperse({ text: ", ", style: normalStyle });
    return SheetUtils.buildRichTexts(elements);
  }

  /// place entries from DataSheet onto empty Schedule
  function placeEntries(): void {
    const entryRange = getEntriesRange();
    const data = entryRange.getRichTextValues();
    // place entries from data sheet into schedule sheet
    DataSheet.forEachEntryGrouped((entries: Entry.IEntry[]) => {
      const first = entries[0];
      if (DateUtils.inRangeInclusive(first.date, dateRange().from, dateRange().until)) {
        let row = entryRow(first.date) - FIRST_ENTRY_ROW;
        let col = entryColumn(first.date, first.location) - FIRST_ENTRY_COLUMN;
        const offset = first.shift.entryDisplayOffset;
        row += offset[0];
        col += offset[1];
        data[row][col] = layoutEntryGroup(entries);
      }
    });
    entryRange.setRichTextValues(data);
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

  /** setup the part of the sheet on the left that lists employees and how much they work. */
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

  /** Setup the sheet and copy the range of entries from the data sheet. */
  export function setup(fDate: Date, tDate: Date): void {
    sheet.clear();
    sheet.setHiddenGridlines(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(FIRST_ENTRY_COLUMN - 2);
    dateRangeCache = undefined;
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
    forEachDayOnSheet((date) => {
      const row = entryRow(date);
      sheet.getRange(row, INDEX_COLUMN).setValue(date).setNumberFormat('ddd", "mmmm" "d');
      sheet.getRange(row, INDEX_COLUMN, 2, 1)
        .mergeVertically()
        .setBorder(true, true, true, true, false, false,  "#000000", SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment("middle");
      Locations.all().forEach((loc) => {
        const col = entryColumn(date, loc);

        // TODO: fix d.ts file for range
        sheet.getRange(row, col, 1, 2).mergeAcross();
        sheet.getRange(row, col, 2, 2)
          .setBorder(true, true, true, true, false, false, "#000000", SpreadsheetApp.BorderStyle.SOLID)
          .setBorder(null, null, null, null, true, true, "#dddddd", SpreadsheetApp.BorderStyle.SOLID);
      });
      if (DateUtils.isWeekend(date)) {
        sheet.getRange(row, INDEX_COLUMN, ROWS_PER_ENTRY, 1 + Locations.all().length * (COLUMNS_PER_ENTRY + 1))
          .setBackground(Config.WEEKEND_COLOR);
      }
    });
    sheet.autoResizeColumn(INDEX_COLUMN);
    placeEntries();
    sheet.getRange("B1").activate();
    // TODO: make validation call is_valid_schedule_entry
    // const rule=SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=IS_VALID_SCHEDULE_ENTRY")
  }

  /** Calls f for each entry on the schedule sheet. */
  export function forEachEntry(f: (schedule: Entry.IEntry) => void): void {
    const data = getEntriesRange().getValues();
    forEachDayOnSheet((date) => {
      Locations.all().forEach((loc) => {
        const row = entryRow(date) - FIRST_ENTRY_ROW;
        const col = entryColumn(date, loc) - FIRST_ENTRY_COLUMN;
        const whole = splitNames(Values.get(data, row, col, Values.asString));
        const firstHalf = splitNames(Values.get(data, row + 1, col, Values.asString));
        const secondHalf = splitNames(Values.get(data, row + 1, col + 1, Values.asString));
        const all =
            [ { shift : Shifts.firstHalf, names : firstHalf },
              { shift : Shifts.secondHalf, names : secondHalf },
              { shift: Shifts.whole, names: whole },
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

  /** Get employees and locations to schedule as dictionary (from the left pane).
   * Returns only those employees who should be placed.
   */
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

  /** Called on edit of a cell. */
  export function onEditCallback(e: GoogleAppsScript.Events.SheetsOnEdit) {
    // check if range is bigger than one cell and if so just recreate the
    // range of the data sheet that is on the schedule.  That is we only
    // try to do minimal work when only single cell was changed.
    // Annoyingly the below turned out not to work (contrary to the docs)
    // as even when I had selected multiple cells NumRows and NumColumns was
    // always 1.  So I get that check (in case this ever gets fixed) but
    // also added a check for the active selection.
    // if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) {
    const activeRange = sheet.getActiveRange();
    if (SheetUtils.isCell(activeRange) && SheetUtils.isCell(e.range)) {
      const range = dateRange();
      const entriesOnSchedule = Prelude.forEachAsList(forEachEntry);
      DataSheet.replaceRange(range.from, range.until, entriesOnSchedule);
      // TODO: redraw everything in this case?
      return;
    }
    // Otherwise do the one cell fast path:
    // figure out which entry was changed, if any
    const entry = cellToEntry(e.range.getRow(), e.range.getColumn());
    // Change wasn't of a entry cell so we are good.
    if (entry !== undefined) {
      // remove any relevant existing entries in the datasheet
      DataSheet.removeMatching(entry.date, entry.location.name, entry.shift.name);
      // and create new ones.
      const employees = splitNames(e.value);
      const entries = employees.map((name: string) => ({ employee: name, ...entry }));
      DataSheet.add(entries);
      // And also redraw that one cell
      // e.range.setRichTextValue(layoutEntryGroup(entries));
      // sheet.getRange(e.range.getRow(), e.range.getColumn()).setValue("TEST");
    }
  }
}

/** @customfunction */
function IS_VALID_SCHEDULE_ENTRY(cell: any) {
  return typeof cell === "string" && ScheduleSheet.validateEntry(cell);
}
