/** @OnlyCurrentDoc */
/** The schedule sheet draws one box per place, each box being subdivided into 3 cells
 * one for each of the standard slots.
 */
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
    const result = entry.split(",").map((e) => e.trim()).filter((e) => e !== "");
    result.sort();
    return result;
  }

  export function validateEntry(entry: string): boolean {
    const employees = EmployeeSheet.byAliasAndHandle();
    return entry.split(",").map((e) => e.trim()).every((e) => {
      const emp = employees[e];
      return emp && emp.employee === e;
    });
  }

  /** return the top row of the box for the given date */
  function dateToRow(date: Date): number {
    return FIRST_ENTRY_ROW + DateUtils.daysBetween(dateRange().from, date) * ROWS_PER_ENTRY;
  }

  /** Convert row number into date. */
  function rowToDate(row: number): Date|undefined {
    if (row < FIRST_ENTRY_ROW) {
      return undefined;
    }
    const dr = dateRange();
    const date = DateUtils.addDays( dr.from, Math.floor((row - FIRST_ENTRY_ROW) / 2) );
    if (!DateUtils.inRangeInclusive(date, dr.from, dr.until)) {
      return undefined;
    }
    return date;
  }

  function placeToColumn({ date, location }: { date: Date; location: Locations.ILocation; }): number {
    return FIRST_ENTRY_COLUMN + location.ndx * (COLUMNS_PER_ENTRY + 1);
  }

  function noteColumn(): number {
    return FIRST_ENTRY_COLUMN + Locations.all().length * (COLUMNS_PER_ENTRY + 1) + 1;
  }

  function cellToSlot(row: number, column: number):
    { date: Date, location: Locations.ILocation, shift: Shifts.IShift } | undefined {
    if (column < FIRST_ENTRY_COLUMN) {
      return undefined;
    }
    const date = rowToDate(row);
    if (!date) { return undefined; }
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
      Logger.log("cellToSlot bug? (row=%s) (column=%s)", row, column);
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

  const boldStyle = SpreadsheetApp.newTextStyle().setBold(true).build();
  const normalStyle = SpreadsheetApp.newTextStyle().build();
  const errorStyle = SpreadsheetApp.newTextStyle().setForegroundColor("red").setBold(true).build();

  function getEmployeeStatus(employee: string): "ok" | "not-in-poll" | "unknown-employee" {
    const employees = EmployeeSheet.byAliasAndHandle();
    const emp = employees[employee];
    if (emp !== undefined && emp.employee === employee) {
      // TODO: write check for poll status
      return "ok";
    } else {
      return "unknown-employee";
    }
  }

  function layoutEntry(entry: Entry.IEntry): GoogleAppsScript.Spreadsheet.RichTextValue {
    const elements =
      entry.employees.map((employee) => {
        const text = employee;
        switch (getEmployeeStatus(text)) {
          case "ok":
            return { text, style: normalStyle };

          case "not-in-poll":
            return { text, style: boldStyle };

          case "unknown-employee":
            return { text, style: errorStyle };
        }
      }).intersperse({ text: ", ", style: normalStyle });
    return SheetUtils.buildRichTexts(elements);
  }

  function slotToCell(slot: Entry.Slot): { row: number, column: number } {
    const row = dateToRow(slot.date);
    const column = placeToColumn(slot);
    const offset = slot.shift.entryDisplayOffset;
    return { row: row + offset[0], column: column + offset[1] };
  }

  /** place entries from DataSheet onto empty Schedule. */
  function placeEntries(): void {
    const entryRange = getEntriesRange();
    const data = entryRange.getRichTextValues();
    // place entries from data sheet into schedule sheet
    DataSheet.forEach((entry: Entry.IEntry) => {
      if (DateUtils.inRangeInclusive(entry.date, dateRange().from, dateRange().until)) {
        const cell = slotToCell(entry);
        Logger.log("placing %s at %s", entry, cell);
        data[cell.row - FIRST_ENTRY_ROW][cell.column - FIRST_ENTRY_COLUMN] = layoutEntry(entry);
      }
    });
    entryRange.setRichTextValues(data);
  }

  export function formulasEmployeeCount(date: Date, loc: Locations.ILocation): string[] {
    function countFormula(cell: string) {
      return "IF(ISBLANK(Schedule!" + cell + ");0;LEN(Schedule!" + cell +
        ")-LEN(SUBSTITUTE(Schedule!" + cell + ';",";""))+1)';
    }
    const row = dateToRow(date);
    const col = placeToColumn({ date, location: loc });
    const wholeDay = countFormula(SheetUtils.a1(row, col));
    const firstHalf = countFormula(SheetUtils.a1(row + 1, col));
    const secondHalf = countFormula(SheetUtils.a1(row + 1, col + 1));
    return ["=" + wholeDay + "+" + firstHalf, "=" + wholeDay + "+" + secondHalf];
  }

  /** setup the part of the sheet on the left that lists employees and how much they work. */
  function setupEmployeeSection(): void {
    const employees = EmployeeSheet.all().map((e) => [ e.employee ]);
    sheet.getRange(FIRST_ENTRY_ROW - 1, 1, 1, 3).setValues([["Mitarbeiter", "Stunden", ""]]).setFontWeight("bold");
    const employeesRange = sheet.getRange(FIRST_ENTRY_ROW, 1, employees.length, 1);
    // const employeeInDoodleRule = SpreadsheetApp.newConditionalFormatRule()
    //   .whenFormulaSatisfied("=ISNA(VLOOKUP(A3; UmfrageAlsTabelle!A2:A; 1; FALSE))")
    //   .setItalic(true)
    //   .setRanges([employeesRange])
    //   .build();
    // sheet.setConditionalFormatRules([employeeInDoodleRule]);
    employeesRange.setValues(employees);
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

  namespace NoteSection {
    function noteRange(): GoogleAppsScript.Spreadsheet.Range {
      const dr = dateRange();
      return sheet.getRange(FIRST_ENTRY_ROW, noteColumn(),
        ROWS_PER_ENTRY * (DateUtils.daysBetween(dr.from, dr.until) + 1), 1);
    }

    export function setup() {
      const dr = dateRange();
      const col = noteColumn();
      noteRange().setBackgroundRGB(255, 255, 153);
      sheet.getRange(1, col).setValue("Notizen").setFontWeight("bold");
      // if Andi starts using notes a lot switch to a single setValues call instead of this loop
      NoteSheet.forEachEntryInRange(dr.from, dr.until, (note) => {
        sheet.getRange(dateToRow(note.date) + note.index, noteColumn()).setValue(note.text);
      });
      sheet.autoResizeColumn(col);
    }

    /** Call f foreach note. */
    export function forEach(f: (note: NoteSheet.Note) => void): void {
      const data = noteRange().getValues();
      data.forEach((row, n) => {
        const date = rowToDate(FIRST_ENTRY_ROW + n)!;
        const firstRowOfDate = dateToRow(date);
        const index = FIRST_ENTRY_ROW + n - firstRowOfDate;
        if (row[0] !== "" && row[0] !== undefined) {
          f({ date, index, text: Values.asString(row[0]) });
        }
      });
    }

    export function save() {
      const notes = Prelude.forEachAsList(forEach);
      const dr = dateRange();
      NoteSheet.replaceRange(dr.from, dr.until, notes);
    }
  }

  /** Setup the sheet and copy the range of entries from the data sheet. */
  export function setup(fDate: Date, tDate: Date): void {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setHiddenGridlines(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(FIRST_ENTRY_COLUMN - 2);
    dateRangeCache = undefined;
    sheet.getRange(1, 1, 1, 4).setValues([["Von", fDate, "bis", tDate]]);
    sheet.getRangeList(["A1", "C1"]).setFontWeight("bold").setHorizontalAlignment("right");
    sheet.getRangeList(["B1", "D1"]).setNumberFormat("yyyy-mm-dd");
    setupEmployeeSection();
    NoteSection.setup();
    // write the column of dates, draw the boxes and setup the formula for #people per entry
    forEachDayOnSheet((date) => {
      const row = dateToRow(date);
      sheet.getRange(row, INDEX_COLUMN).setValue(date).setNumberFormat('ddd", "mmmm" "d');
      sheet.getRange(row, INDEX_COLUMN, 2, 1)
        .mergeVertically()
        .setBorder(true, true, true, true, false, false,  "#000000", SpreadsheetApp.BorderStyle.SOLID)
        .setVerticalAlignment("middle");
      Locations.all().forEach((loc) => {
        const col = placeToColumn({ date, location: loc });

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
    // make sure columns on the right are sized properly
    Locations.all().forEach((loc, ndx) => {
      const col = FIRST_ENTRY_COLUMN + ndx * (COLUMNS_PER_ENTRY + 1);
      sheet.getRange(1, col).setValue(loc.name).setFontWeight("bold");
      sheet.setColumnWidth(col - 1, 10);
      SheetUtils.autoResizeColumns(sheet, col, 2, 80);
    });
    // TODO: make validation call is_valid_schedule_entry
    // const rule=SpreadsheetApp.newDataValidation().requireFormulaSatisfied("=IS_VALID_SCHEDULE_ENTRY")
  }

  /** Calls f for each entry on the schedule sheet. */
  export function forEachEntry(f: (schedule: Entry.IEntry) => void): void {
    const data = getEntriesRange().getValues();
    forEachDayOnSheet((date) => {
      Locations.all().forEach((loc) => {
        const row = dateToRow(date) - FIRST_ENTRY_ROW;
        const col = placeToColumn({ date, location: loc }) - FIRST_ENTRY_COLUMN;
        const whole = splitNames(Values.get(data, row, col, Values.asString));
        const firstHalf = splitNames(Values.get(data, row + 1, col, Values.asString));
        const secondHalf = splitNames(Values.get(data, row + 1, col + 1, Values.asString));
        const all =
            [ { shift : Shifts.firstHalf, names : firstHalf },
              { shift : Shifts.secondHalf, names : secondHalf },
              { shift: Shifts.whole, names: whole },
            ] ;
        all.forEach((e) => {
          if (e.names.length > 0) {
            const entry: Entry.IEntry = {
              date, employees: e.names, location: Locations.all()[loc.ndx], shift: e.shift,
            };
            f(entry);
          }
        });
      });
    });
  }

  const compareSlot =
    Prelude.lexiographic ([
      Prelude.compareBy((s: Entry.Slot) => s.date, DateUtils.compare),
      Prelude.compareBy((s: Entry.Slot) => s.location.name, Prelude.stringCompare),
      Prelude.compareBy((s: Entry.Slot) => s.shift.name, Prelude.stringCompare),
    ]);

  interface Diff extends Entry.Slot {
    employeesData: string[];
    employeesSchedule: string[];
  }

  /** Compare whats on the sheet with what's in the daten section.  Returns a list
   * of all slots that don't match.
   */
  export function compareWithDataSheet(): Diff[] {
    const result = [];
    const dr = dateRange();
    const data =
      Prelude.forEachAsList(DataSheet.forEach, ((e) => DateUtils.inRangeInclusive(e.date, dr.from, dr.until)));
    const schedule =
      Prelude.forEachAsList(forEachEntry);
    let d = 0;
    let s = 0;
    while (d < data.length && s < schedule.length) {
      switch (compareSlot(data[d], schedule[s])) {
        case "lt":
          const diffLt = {
            date: data[d].date,
            location: data[d].location,
            shift: data[d].shift,
            employeesData: data[d].employees,
            employeesSchedule: [],
          };
          result.push(diffLt);
          d++;
          break;

        case "eq":
          if (!Prelude.arrayEqual(data[d].employees, schedule[s].employees)) {
            const diffEq = {
              date: data[d].date,
              location: data[d].location,
              shift: data[d].shift,
              employeesData: data[d].employees,
              employeesSchedule: schedule[s].employees,
            };
            result.push(diffEq);
          }
          d++;
          s++;
          break;

        case "gt":
          const diffGt = {
            date: data[d].date,
            location: data[d].location,
            shift: data[d].shift,
            employeesData: [],
            employeesSchedule: schedule[s].employees,
          };
          result.push(diffGt);
          s++;
          break;
      }
    }
    while (d < data.length) {
      const diffLt = {
        date: data[d].date,
        location: data[d].location,
        shift: data[d].shift,
        employeesData: data[d].employees,
        employeesSchedule: [],
      };
      result.push(diffLt);
    }
    while (d < schedule.length) {
      const diffGt = {
        date: data[d].date,
        location: data[d].location,
        shift: data[d].shift,
        employeesData: [],
        employeesSchedule: schedule[s].employees,
      };
      result.push(diffGt);
    }
    return result;
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

  /** Assume that the schedule is correct and if there are any differences between
   * the ScheduleSheet and the Datasheet, change the DataSheet to match.
   */
  export function syncScheduleToData() {
    const diffs = compareWithDataSheet();
    diffs.forEach((diff) => {
      DataSheet.removeMatching(diff.date, diff.location.name, diff.shift.name);
      const entry = {
        date: diff.date,
        location: diff.location,
        shift: diff.shift,
        employees: diff.employeesSchedule,
      };
      DataSheet.add([entry]);
    });
  }

  function onEditCallbackLogic(e: GoogleAppsScript.Events.SheetsOnEdit): void {
    // In my testing OnEdit events seem to always happen.
    // But on the other hand there are comments on stackexchange indicating
    // that multiple OnEdit events can get coalesced.
    // So to save the changes from the schedule -> data we always use the
    // diff and patch method (aka syncScheduleToData).  This means that if
    // we missed a previous event we will save its effect eventually.
    // We only use the single cell changes to force the redrawing of the
    // slot (so that we can do highlighting of unknown employees etc -- immediately
    // after the change).
    // TODO: handle the notesection in the same way.
    const ev = SheetUtils.onEditEvent(e);
    syncScheduleToData();
    switch (ev.kind) {
      case "mass-change":
        NoteSection.save();
        break;

      case "change":
      case "insert":
      case "clear":
        const slot = cellToSlot(e.range.getRow(), e.range.getColumn());
        if (slot !== undefined) { // change was of a slot
          // Redraw that one slot
          const entry = {
            date: slot.date,
            location: slot.location,
            shift: slot.shift,
            employees: splitNames(ev.value.toString()),
          };
          e.range.setRichTextValue(layoutEntry(entry));
        } else if (e.range.getColumn() === noteColumn()) { // change of a note
          const date = rowToDate(e.range.getRow());
          if (!date) {
            /// notes outside the date column are ignored
            return;
          }
          const firstRowOfDate = dateToRow(date);
          const ndx = e.range.getRow() - firstRowOfDate;
          switch (ev.kind) {
            case "clear":
              NoteSheet.deleteMatching(date, ndx);
              break;
            case "change":
            case  "insert":
              NoteSheet.addOrReplace({ date, index: ndx, text: e.value });
              break;
          }
        }
        break;
    }
  }

  let insideOnEditCallback = false;

  export function onEditCallback(e: GoogleAppsScript.Events.SheetsOnEdit): void {
    if (!insideOnEditCallback) {
      Logger.log("--> OnEditCallback");
      insideOnEditCallback = true;
      try {
        onEditCallbackLogic(e);
      } catch (error) {
        // do nothing
      }
      insideOnEditCallback = false;
      Logger.log("<-- OnEditCallback");
    } else {
      Logger.log("Recursive OnEditCallback -- not doing anything");
    }
  }
}

function testCompare() {
  const result = ScheduleSheet.compareWithDataSheet();
  Logger.log("diffs: %s", result.length);
  for (let i = 0; i < Math.min(result.length, 4); i++) {
    Logger.log("diff: %s", result[i]);
  }
}
