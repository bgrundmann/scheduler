/** @OnlyCurrentDoc */
// TODO:
// - A way to get list of problems (aka where we have scheduled someone who hasnt doodled)
// - More summaries of days worked etc (sparkline?)
// - Ability to rerun doodle
// - Dedup entries in poll

function flatten<T>(a: T[][]): T[] {
  const empty: T[] = [];
  return empty.concat(...a);
}

function findIndex<T>(a: T[], pred: (elem: T) => boolean): number | undefined {
  const l = a.length;
  for (let ndx = 0; ndx < l; ndx++) {
    if (pred(a[ndx])) {
      return ndx;
    }
  }
  return undefined;
}

function find<T>(a: T[], pred: (elem: T) => boolean): T | undefined {
  const ndx = findIndex(a, pred);
  if (ndx !== undefined) {
    return a[ndx];
  }
  return undefined;
}

function assertDate(v: unknown): Date {
  if (v instanceof Date) {
    return v;
  }
  throw Error(`Expected a date, got ${v}`);
}

function cellAsString(v: unknown): string {
  if (typeof v === "string") {
    return v;
  }
  return String(v);
}

namespace Shift {
  export interface T {
    start: number;
    stop: number;
  }

  export function equal(s1: T, s2: T): boolean {
    return s1.start === s2.start && s1.stop === s2.stop;
  }
}

namespace DateUtils {
  export function copy(d: Date): Date {
    return new Date(d.getTime());
  }
  export function nextDay(d: Date): Date {
    const c = copy(d);
    c.setDate(d.getDate() + 1);
    return c;
  }
  /// Add the given number of days
  export function addDays(d: Date, n: number): Date {
    const c = copy(d);
    c.setDate(d.getDate() + n);
    return c;
  }
  export function isWeekend(d: Date): boolean {
    const w = d.getDay();
    return w === 6 || w === 0;
  }
  /// Monday is 0, Tuesday 1, ... Sunday is 6
  export function dayOfWeekStartingMonday(d: Date): number {
    const w = d.getDay();
    if (w === 0) {
      return 6;
    } else {
      return w - 1;
    }
  }
  function truncToDay(d: Date): Date {
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  /** Compare the passed in Date objects as Dates (independently of timezone and
   * ignoring anything with finer granularity than a day).
   */
  export function inRangeInclusive(d: Date, low: Date, upp: Date): boolean {
    const dt = truncToDay(d);
    const lt = truncToDay(low);
    const ut = truncToDay(upp);
    return lt.getTime() <= dt.getTime() && dt.getTime() <= ut.getTime();
  }
  /** Compare the passed in Date objects as Dates (independently of timezone
   * and ignoring anything with finer granularity than a day).
   */
  export function equal(d1: Date, d2: Date): boolean {
    return truncToDay(d1).getTime() === truncToDay(d2).getTime();
  }

  /** Call f once for each day in the range lower - upper (inclusive).
   */
  export function forEachDay(
    lower: Date,
    upper: Date,
    f: (d: Date, counter: number) => void
  ): void {
    let d = lower;
    let n = 0;
    while (d <= upper) {
      f(d, n);
      n++;
      d = nextDay(d);
    }
  }
  export function mondayStartingWeekContaining(d: Date): Date {
    const res = copy(d);
    while (res.getDay() !== 1) {
      res.setDate(res.getDate() - 1);
    }
    return res;
  }
  export function diff(bigger: Date, smaller: Date): number {
    return Math.round(
      (bigger.getTime() - smaller.getTime()) / (1000 * 60 * 60 * 24)
    );
  }
  const isoRegex = /\d{4}-\d{2}-\d{2}/;

  /*** Parse a date in extended iso format */
  export function parseISODate(s: string): Date | undefined {
    const r = isoRegex.exec(s);
    if (r !== null) {
      return new Date(s);
    }
    return undefined;
  }

  export function toISODate(d: Date): string {
    const s = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    return Utilities.formatDate(d, s, "YYYY-MM-dd");
  }

  export function max(a: Date, b: Date): Date {
    if (a.getTime() >= b.getTime()) {
      return a;
    }
    return b;
  }

  export function min(a: Date, b: Date): Date {
    if (a.getTime() <= b.getTime()) {
      return a;
    }
    return b;
  }
}

namespace DoodleParser {
  // TODO: parse the year from the sheet
  const year = 2019;

  const germanMonthNames = (() => {
    const res: Record<string, number> = {};
    [
      "Januar",
      "Februar",
      "März",
      "April",
      "Mai",
      "Juni",
      "Juli",
      "August",
      "September",
      "Oktober",
      "November",
      "Dezember",
    ].forEach((name, ndx) => {
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
  function parseMergedRow(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    row: number,
    column: number,
    lastColumn: number
  ): ICell[] {
    const result = [];
    const ranges = sheet
      .getRange(row, column, 1, lastColumn - column)
      .getMergedRanges();
    ranges.sort((a, b) => a.getColumn() - b.getColumn());
    let last = column - 1;
    if (ranges.length === 0) {
      for (let c = last + 1; c <= lastColumn; c++) {
        result.push({
          column: c,
          lastColumn: c,
          value: sheet.getRange(row, c).getValue(),
        });
      }
    }
    // TODO: handle single cells at beginning or end
    ranges.forEach((r) => {
      // the merged cells aren't in the array returned by getMergedRanges, so
      // we need to get them in another way.
      for (let c = last + 1; c < r.getColumn(); c++) {
        result.push({
          column: c,
          lastColumn: c,
          value: sheet.getRange(row, c).getValue(),
        });
      }
      result.push({
        column: r.getColumn(),
        lastColumn: r.getLastColumn(),
        value: r.getValue(),
      });
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

  function forEachShift(
    monthAndYears: any[],
    days: any[],
    times: any[],
    f: (shift: IShift, column: number) => void
  ): void {
    let month = 0;
    let day = 0;
    let time = 0;
    const lastColumn = monthAndYears[monthAndYears.length - 1].lastColumn;

    for (let column = monthAndYears[0].column; column <= lastColumn; column++) {
      if (
        !betweenInclusive(
          column,
          monthAndYears[month].column,
          monthAndYears[month].lastColumn
        )
      ) {
        month++;
      }
      if (!betweenInclusive(column, days[day].column, days[day].lastColumn)) {
        day++;
      }
      if (
        !betweenInclusive(column, times[time].column, times[time].lastColumn)
      ) {
        time++;
      }
      let r = monthAndYearRegex.exec(monthAndYears[month].value);
      if (r == null) {
        throw Error("Failed to parse month and year in doodle");
      }
      const monthValue = germanMonthNames[r[1]];
      const year2 = Number(r[2]);

      r = dayRegex.exec(days[day].value);
      if (r == null) {
        throw Error("Failed to parse day in doodle");
      }
      const dayValue = Number(r[2]);

      // Parser would be more robust if this lookup failed with an error message as well
      r = timeRegex.exec(times[time].value);
      if (r == null) {
        throw Error("Failed to parse time in doodle");
      }
      const timeStartHour = Number(r[1]);
      const timeStartMinute = Number(r[2]);
      const timeEndHour = Number(r[3]);
      const timeEndMinute = Number(r[4]);
      f(
        {
          year: year2,
          month: monthValue,
          day: dayValue,
          timeStartHour,
          timeStartMinute,
          timeEndHour,
          timeEndMinute,
        },
        column
      );
    }
  }

  export interface Entry {
    employee: string;
    date: Date;
    /** minutes since beginning of day. */
    start: number;
    /** minutes since beginning of day. */
    stop: number;
  }

  /** Parse the doodle sheet (which must be called Umfrage) and return a list of entries found.
   * Throws an error if it couldn't parse an entry.
   */
  export function parse(): Entry[] {
    const ss = SpreadsheetApp.getActive();
    const doodle = ss.getSheetByName("Umfrage");
    const employeeDict = EmployeeSheet.get();
    const monthAndYears = parseMergedRow(doodle, 4, 2, doodle.getLastColumn());
    const days = parseMergedRow(doodle, 5, 2, doodle.getLastColumn());
    const times = parseMergedRow(doodle, 6, 2, doodle.getLastColumn());

    const valuesAndMore = doodle
      .getRange(7, 1, doodle.getLastRow() - 6, doodle.getLastColumn() - 1)
      .getValues();
    const summaryRow = findIndex(valuesAndMore, (row) => row[0] === "Anzahl");
    if (summaryRow === undefined) {
      throw Error("Konnte die Anzahl Zeile nicht finden");
    }
    const values = valuesAndMore.slice(undefined, summaryRow);
    const result: Entry[] = [];
    forEachShift(monthAndYears, days, times, (d, c) => {
      // NOTE: forEachShift column is absolute column on sheet (with 1 being A)
      // But values array is 0 based
      for (const row of values) {
        const parsedName = row[0].toString();
        if (!(parsedName in employeeDict)) {
          throw Error(`Unbekannter Mitarbeiter ${parsedName}`);
        }
        const employee = employeeDict[parsedName].handle;
        const ok = row[c - 1] === "OK";
        if (ok) {
          const date = new Date(d.year, d.month, d.day);
          const start = d.timeStartHour * 60 + d.timeStartMinute;
          const stop = d.timeEndHour * 60 + d.timeEndMinute;
          result.push({ employee, date, start, stop });
        }
      }
    });
    return result;
  }
}

namespace SlotParser {
  /** A slot contains a list of items. */
  export type Item =
    | { kind: "default"; name: string }
    | {
        kind: "specified";
        name: string;
        start: number;
        stop: number;
        duration: number;
      };

  const shortOfDayRegex = RegExp("^[0-9]{1,2}$");
  const ofDayRegex = RegExp("^([0-9]{1,2}):([0-9]{2})$");

  /** Parse a time of day and return it as minutes since midnight.
   * Understood formats:
   * hh (0 - 23, with optional leading 0)
   * hh:mm
   *
   * Returns undefined if the string couldn't be parsed.
   */
  function parseTimeOfDay(s: string): number | undefined {
    const m1 = shortOfDayRegex.exec(s);
    if (m1 !== null) {
      return Number(s) * 60;
    }
    const m2 = ofDayRegex.exec(s);
    if (m2 !== null) {
      return Number(m2[1]) * 60 + Number(m2[2]);
    }
    return undefined;
  }

  /** Parse a time range and return start time, stop time (in minutes since midnight) and
   * duration (in minutes).
   */
  export function parseTimeRange(
    s: string
  ): { start: number; stop: number; duration: number } | undefined {
    const parts = s.split("-");
    if (parts.length !== 2) {
      return undefined;
    }
    const start = parseTimeOfDay(parts[0]);
    const stop = parseTimeOfDay(parts[1]);
    if (start !== undefined && stop !== undefined) {
      return { start, stop, duration: stop - start };
    }
    return undefined;
  }

  function parseItem(s: string): Item {
    const m = s.split(" ");
    if (m.length === 2) {
      const range = parseTimeRange(m[1]);
      if (range === undefined) {
        throw Error(`Verstehe die angegebene Zeitspanne ${m[1]} nicht`);
      }
      return {
        kind: "specified",
        name: m[0],
        ...range,
      };
    } else {
      return { kind: "default", name: s };
    }
  }

  const commentRegex = /(^|[^0-9])-[^0-9]/;

  /** Parse the value of a slot. Throws an error if the input is invalid. */
  export function parse(slot: string): Item[] {
    const startOfComment = slot.search(commentRegex);
    let slotWithoutComment = "";
    if (startOfComment === -1) {
      slotWithoutComment = slot;
    } else {
      slotWithoutComment = slot.slice(undefined, startOfComment);
    }
    const employees = slotWithoutComment
      .split(",")
      .map((s) => s.trim())
      .filter((s) => s !== "");
    return employees.map((e) => parseItem(e));
  }

  /** Given the textual representation of a slot remove all occurrences of any
   * of the given employees.
   * TODO: Deal with names in comments.
   */
  export function removeEmployees(slot: string, employees: string[]): string {
    const names = employees.join("|");
    const interestingItemRegex = RegExp(
      `(\\b(${names})\\b( +[0-9-:]+)? *,?)`,
      "g"
    );
    const res = slot.replace(interestingItemRegex, "").trim();
    if (res.length > 0 && res[res.length - 1] === ",") {
      return res.substr(0, res.length - 1);
    } else {
      return res;
    }
  }
}

namespace SheetUtils {
  export function autoResizeColumns(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    startColumn: number,
    numColumns: number,
    minWidthPixels?: number
  ) {
    sheet.autoResizeColumns(startColumn, numColumns);
    if (minWidthPixels) {
      for (let c = startColumn; c < startColumn + numColumns; c++) {
        if (sheet.getColumnWidth(c) < minWidthPixels) {
          sheet.setColumnWidth(c, minWidthPixels);
        }
      }
    }
  }
}

namespace Locations {
  export interface Location {
    name: string;
    ndx: number;
  }

  const theList: Location[] = [
    { name: "Marktgasse" },
    { name: "Ammergasse" },
    { name: "Buero" },
    { name: "Online" },
    { name: "Extras" },
  ].map((e, ndx) => ({ ndx, ...e }));

  const planner: Location[] = [{ name: "Unverplant", ndx: theList.length }];
  const doodle: Location[] = [{ name: "Doodle", ndx: theList.length + 1 }];

  const theListWithPlanner: Location[] = theList.concat(planner);
  const theListWithPlannerAndDoodle: Location[] = theListWithPlanner.concat(
    doodle
  );

  export function all(
    mode:
      | "include-planner-and-doodle"
      | "exclude-planner-and-doodle"
      | "include-planner-exclude-doodle"
      | "only-doodle"
      | "only-planner"
  ): Location[] {
    switch (mode) {
      case "include-planner-exclude-doodle":
        return theListWithPlanner;
      case "include-planner-and-doodle":
        return theListWithPlannerAndDoodle;
      case "exclude-planner-and-doodle":
        return theList;
      case "only-doodle":
        return doodle;
      case "only-planner":
        return planner;
    }
  }
}

namespace EmployeeSheet {
  function getSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    return SpreadsheetApp.getActive().getSheetByName("Mitarbeiter");
  }

  /** Copy list of employees from employees sheet to the given Range.
   * Returns the number of employees.
   */
  export function copyEmployees(
    dst: GoogleAppsScript.Spreadsheet.Range,
    transposed?: "transposed"
  ): number {
    const trueIfTransposed = transposed !== undefined;
    const employeeSheet = getSheet();
    // TODO: check if copyTo is the function to use.
    const src = employeeSheet.getRange(2, 1, employeeSheet.getLastRow(), 1);
    src.copyTo(
      dst,
      SpreadsheetApp.CopyPasteType.PASTE_VALUES,
      trueIfTransposed
    );
    const numEmployees = src.getNumRows();
    return numEmployees;
  }

  export interface Employee {
    handle: string; // aka "andi"
    alias?: string; // aka "Der Boss"
  }

  /** Return a dictionary from both names and aliases to */
  export function get(): Record<string, Employee> {
    const employeeSheet = getSheet();
    const src = employeeSheet
      .getRange(2, 1, employeeSheet.getLastRow(), 2)
      .getValues();
    const res: Record<string, Employee> = {};
    src.forEach((row) => {
      const handle = String(row[0]);
      const alias = row[1] !== "" ? String(row[1]) : undefined;
      const employee = { handle, alias };
      if (alias !== undefined) {
        res[alias] = employee;
      }
      res[handle] = employee;
    });
    return res;
  }
}

/** Responsible for the initial setup of all the sheets */
namespace SheetLayouter {
  const EMPLOYEE_COLUMNS = 3;
  const INDEX_COLUMN = EMPLOYEE_COLUMNS + 2;
  const FIRST_ENTRY_COLUMN = INDEX_COLUMN + 2;
  const FIRST_ENTRY_ROW = 5;
  const ROWS_PER_ENTRY = 2;
  const COLUMNS_PER_ENTRY = 2;
  const FROM_DATE_ROW = 1;
  const UNTIL_DATE_ROW = FROM_DATE_ROW + 1;
  const DATE_COLUMN = FIRST_ENTRY_COLUMN;
  const WEEKEND_COLOR = "#FFF2CC";
  const DOODLE_COLOR = "#5B95F9";

  function columnLetter(n: number): string {
    const res = [];
    while (n > 0) {
      const digit = n % 26;
      if (digit === 0) {
        res.push("Z");
        n = Math.floor(n / 26 - 1);
      } else {
        res.push(String.fromCharCode("A".charCodeAt(0) + digit - 1));
        n = Math.floor(n / 26);
      }
    }
    res.reverse();
    return res.join("");
  }

  function a1(row: number, column: number, sheetName?: string): string {
    const cell = columnLetter(column) + String(row);
    if (sheetName !== undefined) {
      return "'" + sheetName + "'!" + cell;
    }
    return cell;
  }

  function a1Range(
    row1: number,
    column1: number,
    row2: number,
    column2: number,
    sheetName?: string
  ): string {
    return a1(row1, column1, sheetName) + ":" + a1(row2, column2, sheetName);
  }

  function getReferencesToSlotsForNthDay(
    scheduleSheetName: string,
    n: number,
    mode: "exclude-planner-and-doodle" | "only-doodle"
  ): string {
    return Locations.all(mode)
      .map((loc) => {
        const whole = a1(
          FIRST_ENTRY_ROW + n * ROWS_PER_ENTRY,
          FIRST_ENTRY_COLUMN + loc.ndx * COLUMNS_PER_ENTRY,
          scheduleSheetName
        );
        const firstHalf = a1(
          FIRST_ENTRY_ROW + n * ROWS_PER_ENTRY + 1,
          FIRST_ENTRY_COLUMN + loc.ndx * COLUMNS_PER_ENTRY,
          scheduleSheetName
        );
        const secondHalf = a1(
          FIRST_ENTRY_ROW + n * ROWS_PER_ENTRY + 1,
          FIRST_ENTRY_COLUMN + loc.ndx * COLUMNS_PER_ENTRY + 1,
          scheduleSheetName
        );
        return `${whole};${firstHalf};${secondHalf}`;
      })
      .join(";");
  }

  function countEmployeesFormula(slot: string): string {
    return `IF(ISBLANK(${slot});0;LEN(${slot})-LEN(REGEXREPLACE(${slot};","; ""))+1)`;
  }

  /** Count number of employees available in the first and second half of the day as per doodle */
  function countAvailableEmployeesFormula(
    nthDay: number
  ): { firstHalf: string; secondHalf: string } {
    const doodle = Locations.all("only-doodle")[0];
    const whole = a1(
      FIRST_ENTRY_ROW + nthDay * ROWS_PER_ENTRY,
      FIRST_ENTRY_COLUMN + doodle.ndx * COLUMNS_PER_ENTRY
    );
    const firstHalf = a1(
      FIRST_ENTRY_ROW + nthDay * ROWS_PER_ENTRY + 1,
      FIRST_ENTRY_COLUMN + doodle.ndx * COLUMNS_PER_ENTRY
    );
    const secondHalf = a1(
      FIRST_ENTRY_ROW + nthDay * ROWS_PER_ENTRY + 1,
      FIRST_ENTRY_COLUMN + doodle.ndx * COLUMNS_PER_ENTRY + 1
    );
    const wholeFormula = countEmployeesFormula(whole);
    return {
      firstHalf: `=${countEmployeesFormula(firstHalf)}+${wholeFormula}`,
      secondHalf: `=${countEmployeesFormula(secondHalf)}+${wholeFormula}`,
    };
  }

  /** Return the range of dates displayed. */
  function getDates(
    sheet: GoogleAppsScript.Spreadsheet.Sheet
  ): { from: Date; until: Date } {
    const [[fromV], [untilV]] = sheet
      .getRange(FROM_DATE_ROW, DATE_COLUMN, 2, 1)
      .getValues();
    const from = assertDate(fromV);
    const until = assertDate(untilV);
    return { from, until };
  }

  function getDefaultTimeRanges(
    sheet: GoogleAppsScript.Spreadsheet.Sheet
  ): {
    whole: { start: number; stop: number; duration: number };
    firstHalf: { start: number; stop: number; duration: number };
    secondHalf: { start: number; stop: number; duration: number };
  } {
    const [[wholeCell], [first], [second]] = sheet
      .getRange(1, 2, 3, 1)
      .getValues();
    const whole = SlotParser.parseTimeRange(String(wholeCell))!;
    whole.duration = whole.duration - 60;
    return {
      whole,
      firstHalf: SlotParser.parseTimeRange(String(first))!,
      secondHalf: SlotParser.parseTimeRange(String(second))!,
    };
  }

  const saturdayTimeRange = {
    start: 9 * 60 + 45,
    stop: 16 * 60,
    duration: 6 * 60 + 15,
  };

  const sundayTimeRange = {
    start: 13 * 60,
    stop: 18 * 60,
    duration: 5 * 60 * 1.5,
  };

  function setupComputationSheet(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    fromDate: Date,
    toDate: Date,
    mode: "exclude-planner-and-doodle" | "only-doodle"
  ): void {
    const scheduleSheetName = scheduleSheet.getName();
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setFrozenRows(2);
    sheet.setFrozenColumns(1);
    const numDays = DateUtils.diff(toDate, fromDate) + 1;
    const numEmployees = EmployeeSheet.copyEmployees(
      sheet.getRange(1, 2),
      "transposed"
    );
    sheet.getRange(1, 2, 1, numEmployees).setFontWeight("bold");
    // Range starting at the first row of data (aka row[0] = row of fromDate)
    const theRange = sheet.getRange(
      2,
      1,
      numDays + 1, // +1 for the row of sums
      numEmployees + 1
    );
    const cells = theRange.getValues();
    const employees = a1(1, 2) + ":" + a1(1, numEmployees + 2 - 1);
    const defaults = getDefaultTimeRanges(scheduleSheet);
    const times = `${defaults.whole.duration};${defaults.firstHalf.duration};${
      defaults.secondHalf.duration
    }`;
    const saturdayTimes = `${saturdayTimeRange.duration};${
      saturdayTimeRange.duration
    };${saturdayTimeRange.duration}`;
    const sundayTimes = `${sundayTimeRange.duration};${
      sundayTimeRange.duration
    };${sundayTimeRange.duration}`;
    DateUtils.forEachDay(fromDate, toDate, (date, nth) => {
      const row = nth + 1; // to skip the sum row
      const slots = getReferencesToSlotsForNthDay(scheduleSheetName, nth, mode);
      let timesWeekdaySpecific = "";
      const wd = date.getDay();
      if (wd === 6) {
        timesWeekdaySpecific = saturdayTimes;
      } else if (wd === 0) {
        timesWeekdaySpecific = sundayTimes;
      } else {
        timesWeekdaySpecific = times;
      }
      cells[row][0] = date;
      cells[
        row
      ][1] = `=SCHEDULECOUNT(${employees};${timesWeekdaySpecific};${slots})`;
    });
    for (let e = 0; e < numEmployees; e++) {
      const thatColumn = a1Range(3, 2 + e, 3 + numDays - 1, 2 + e);
      cells[0][e + 1] = `=SUM(${thatColumn})`;
    }
    theRange.setValues(cells);
  }

  export interface Slot {
    date: Date;
    location: number;
    shiftKind: number;
    row: number;
    column: number;
  }

  /** Return the range of all the entries for the given day. Result is undefined if the passed
   * in date is not on the sheet.
   */
  export function rangeOfEntriesOfDay(
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    date: Date,
    mode: "include-planner-exclude-doodle"
  ): {
    range: GoogleAppsScript.Spreadsheet.Range;
    row: number;
    column: number;
  } {
    const dates = getDates(scheduleSheet);
    const offset = DateUtils.diff(date, dates.from) * ROWS_PER_ENTRY;
    const row = FIRST_ENTRY_ROW + offset;
    const columns = Locations.all(mode).length * COLUMNS_PER_ENTRY;
    return {
      range: scheduleSheet.getRange(
        row,
        FIRST_ENTRY_COLUMN,
        ROWS_PER_ENTRY,
        columns
      ),
      row,
      column: FIRST_ENTRY_COLUMN,
    };
  }

  /** Convert the position of a cell to the address of the relevant slot, if there
   * is one at that position.
   */
  export function cellToSlot(
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    mode: "include-planner-exclude-doodle",
    cell: { row: number; column: number }
  ): undefined | Slot {
    const dates = getDates(scheduleSheet);
    const numDays = DateUtils.diff(dates.until, dates.from) + 1;
    const rows = numDays * ROWS_PER_ENTRY;
    const columns = Locations.all(mode).length * COLUMNS_PER_ENTRY;
    if (
      FIRST_ENTRY_ROW <= cell.row &&
      cell.row < FIRST_ENTRY_ROW + rows &&
      FIRST_ENTRY_COLUMN <= cell.column &&
      cell.column < FIRST_ENTRY_COLUMN + columns
    ) {
      const date = DateUtils.addDays(
        dates.from,
        Math.floor((cell.row - FIRST_ENTRY_ROW) / ROWS_PER_ENTRY)
      );
      const row0or1 = (cell.row - FIRST_ENTRY_ROW) % 2;
      let shiftKind = 0;
      if (row0or1 === 1) {
        if ((cell.column - FIRST_ENTRY_COLUMN) % COLUMNS_PER_ENTRY === 0) {
          shiftKind = 1;
        } else {
          shiftKind = 2;
        }
      }
      const location = Math.floor(
        (cell.column - FIRST_ENTRY_COLUMN) / COLUMNS_PER_ENTRY
      );
      return { date, location, shiftKind, ...cell };
    }
    return undefined;
  }

  /** (Re-)create the formulas on the sheets.  Called both initially
   * when the sheets are first populated and also on user demand when
   * the user updated the employee list.  For the latter reason it must
   * not touch any user data on the sheets.
   * The date range passed in MUST match the date range on the schedule
   * sheet.
   */
  export function updateFormulas({
    scheduleSheet,
    doodleSheet,
    workSheet,
    fromDate,
    toDate,
  }: {
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet;
    doodleSheet: GoogleAppsScript.Spreadsheet.Sheet;
    workSheet: GoogleAppsScript.Spreadsheet.Sheet;
    fromDate: Date;
    toDate: Date;
  }): void {
    // Setup the sheet summing up the working minutes
    setupComputationSheet(
      workSheet,
      scheduleSheet,
      fromDate,
      toDate,
      "exclude-planner-and-doodle"
    );
    // Setup the sheet summing up the doodled minutes
    setupComputationSheet(
      doodleSheet,
      scheduleSheet,
      fromDate,
      toDate,
      "only-doodle"
    );
  }

  /** The initial setup of a sheet.  Intended to be called exactly once for each sheet
   * (immediately after creation of the sheet).
   */
  export function setup(
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    workSheet: GoogleAppsScript.Spreadsheet.Sheet,
    doodleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    fromDate: Date,
    toDate: Date
  ): void {
    function columnOfEntry(location: Locations.Location) {
      return FIRST_ENTRY_COLUMN + location.ndx * COLUMNS_PER_ENTRY;
    }
    scheduleSheet.clear();
    scheduleSheet.clearConditionalFormatRules();
    scheduleSheet.setHiddenGridlines(true);
    scheduleSheet.setFrozenRows(FIRST_ENTRY_ROW - 2);
    scheduleSheet.setFrozenColumns(INDEX_COLUMN + 1);
    scheduleSheet.setColumnWidth(INDEX_COLUMN - 1, 5);
    scheduleSheet.setRowHeight(FIRST_ENTRY_ROW - 1, 5);
    scheduleSheet
      .getRange(1, 1, 3, 2)
      .setValues([
        ["Ganztags", "9:45-19:00"],
        ["Vormittags", "9:45-14:00"],
        ["Nachmittags", "13-19"],
      ]);
    scheduleSheet.getRange(1, 1, 3, 1).setFontWeight("bold");
    // setup pane on the left that shows computations.
    scheduleSheet
      .getRange(FIRST_ENTRY_ROW, 1, 1, 2)
      .setValues([["Mitarbeiter", "h"]])
      .setFontWeight("bold");
    const numEmployees = EmployeeSheet.copyEmployees(
      scheduleSheet.getRange(FIRST_ENTRY_ROW + 1, 1)
    );
    const a = [];
    const workSheetName = workSheet.getName();
    for (let e = 0; e < numEmployees; e++) {
      a[e] = [`=${a1(2, 2 + e, workSheetName)}/60`];
    }
    scheduleSheet
      .getRange(FIRST_ENTRY_ROW + 1, 2, numEmployees, 1)
      .setValues(a);
    scheduleSheet
      .getRange(FIRST_ENTRY_ROW, 1, numEmployees + 1, 2)
      .applyRowBanding(SpreadsheetApp.BandingTheme.BLUE);
    // Setup the main scheduling section
    DateUtils.forEachDay(fromDate, toDate, (date, nth) => {
      const row = FIRST_ENTRY_ROW + nth * ROWS_PER_ENTRY;
      // Boxes around the dates in the index column
      scheduleSheet
        .getRange(row, INDEX_COLUMN)
        .setValue(date)
        .setNumberFormat('ddd", "mmmm" "d');
      scheduleSheet
        .getRange(row, INDEX_COLUMN, 1, 2)
        .mergeAcross()
        .setHorizontalAlignment("center");
      const f = countAvailableEmployeesFormula(nth);
      scheduleSheet
        .getRange(row + 1, INDEX_COLUMN, 1, 2)
        .setValues([[f.firstHalf, f.secondHalf]]);
      scheduleSheet
        .getRange(row, INDEX_COLUMN, 2, 2)
        .setBorder(
          true,
          true,
          true,
          true,
          false,
          false,
          "#000000",
          SpreadsheetApp.BorderStyle.SOLID
        );
      Locations.all("include-planner-and-doodle").forEach((loc) => {
        const col = columnOfEntry(loc);

        scheduleSheet.getRange(row, col, 1, 2).mergeAcross();
        scheduleSheet
          .getRange(row, col, 2, 2)
          .setBorder(
            true,
            true,
            true,
            true,
            false,
            false,
            "#000000",
            SpreadsheetApp.BorderStyle.SOLID
          )
          .setBorder(
            null,
            null,
            null,
            null,
            true,
            true,
            "#dddddd",
            SpreadsheetApp.BorderStyle.SOLID
          );
      });
      if (DateUtils.isWeekend(date)) {
        scheduleSheet
          .getRange(
            row,
            INDEX_COLUMN,
            ROWS_PER_ENTRY,
            2 +
              Locations.all("include-planner-and-doodle").length *
                COLUMNS_PER_ENTRY
          )
          .setBackground(WEEKEND_COLOR);
      }
    });
    scheduleSheet.autoResizeColumns(INDEX_COLUMN, 2);
    const headers: string[] = flatten(
      Locations.all("include-planner-and-doodle").map((loc) => [loc.name, ""])
    );
    scheduleSheet
      .getRange(
        FIRST_ENTRY_ROW - 2,
        INDEX_COLUMN,
        1,
        Locations.all("include-planner-and-doodle").length * COLUMNS_PER_ENTRY +
          2
      )
      .setValues([["", ""].concat(headers)])
      .setBackground("#F7CB4D")
      .setFontWeight("bold");
    scheduleSheet
      .getRange(
        FIRST_ENTRY_ROW - 2,
        FIRST_ENTRY_COLUMN +
          Locations.all("only-doodle")[0].ndx * COLUMNS_PER_ENTRY,
        1,
        COLUMNS_PER_ENTRY
      )
      .setBackground(DOODLE_COLOR);
    SheetUtils.autoResizeColumns(
      scheduleSheet,
      FIRST_ENTRY_COLUMN,
      COLUMNS_PER_ENTRY * Locations.all("include-planner-and-doodle").length,
      160
    );
    scheduleSheet.autoResizeRows(FIRST_ENTRY_ROW - 2, 1);
    updateFormulas({ scheduleSheet, doodleSheet, workSheet, fromDate, toDate });
    scheduleSheet
      .getRange(FIRST_ENTRY_ROW - 4, INDEX_COLUMN + 1, 2, 2)
      .setValues([["Von", fromDate], ["Bis", toDate]]);
    scheduleSheet
      .getRange(FIRST_ENTRY_ROW - 4, INDEX_COLUMN + 1, 2, 1)
      .setFontWeight("bold")
      .setHorizontalAlignment("right");
    // dateRangeCache = undefined;
    // sheet.getRange(1, 1, 1, 4).setValues([["Von", fDate, "bis", tDate]]);
    // sheet
    //   .getRangeList(["A1", "C1"])
    //   .setFontWeight("bold")
    //   .setHorizontalAlignment("right");
    // sheet.getRangeList(["B1", "D1"]).setNumberFormat("yyyy-mm-dd");
  }

  export function replaceDoodle(
    scheduleSheet: GoogleAppsScript.Spreadsheet.Sheet,
    entries: DoodleParser.Entry[]
  ) {
    // TODO: Make it such that one can run this function multiple
    // times (in particular after planning has started) and it should
    // do the right thing (TM).
    // TODO: check that sheet mentions what the default times are...
    const doodle = Locations.all("only-doodle")[0];
    const { from, until } = getDates(scheduleSheet);
    const numDays = DateUtils.diff(until, from) + 1;
    const defaults = getDefaultTimeRanges(scheduleSheet);
    const range = scheduleSheet.getRange(
      FIRST_ENTRY_ROW,
      FIRST_ENTRY_COLUMN + doodle.ndx * COLUMNS_PER_ENTRY,
      numDays * ROWS_PER_ENTRY,
      COLUMNS_PER_ENTRY
    );
    const values = range.getValues();
    // Clear content if any
    values.forEach((row) => {
      for (let i = 0; i < row.length; i++) {
        row[i] = "";
      }
    });
    entries.forEach((entry) => {
      const row = DateUtils.diff(entry.date, from) * ROWS_PER_ENTRY;
      function add(rowOff: number, colOff: number) {
        const old = values[row + rowOff][colOff];
        values[row + rowOff][colOff] =
          old === "" ? entry.employee : old + ", " + entry.employee;
      }
      if (Shift.equal(entry, defaults.whole)) {
        add(0, 0);
      } else if (Shift.equal(entry, defaults.firstHalf)) {
        add(1, 0);
      } else if (Shift.equal(entry, defaults.secondHalf)) {
        add(1, 1);
      } else {
        throw Error(`Huh? -- Don't know where to place this ${entry}`);
      }
    });
    range.setValues(values);
    const planner = Locations.all("only-planner")[0];
    const plannerRange = scheduleSheet.getRange(
      FIRST_ENTRY_ROW,
      FIRST_ENTRY_COLUMN + planner.ndx * COLUMNS_PER_ENTRY,
      numDays * ROWS_PER_ENTRY,
      COLUMNS_PER_ENTRY
    );
    plannerRange.setValues(values);
  }
}

namespace SheetsManager {
  type Sheet = GoogleAppsScript.Spreadsheet.Sheet;
  export interface Item {
    /** The date range covered by this set of sheets. */
    from: Date;
    until: Date;
    /** this contains the schedule and is the main UI element. */
    scheduleSheet: Sheet;
    /** This contains the result of the SCHEDULECOUNT function for the regular slots */
    workSheet: Sheet;
    /** This contains the result of the SCHEDULECOUNT function for the doodle slots */
    doodleSheet: Sheet;
  }

  interface ItemInProgress {
    from: Date;
    until: Date;
    scheduleSheet?: Sheet;
    workSheet?: Sheet;
    doodleSheet?: Sheet;
  }

  const nameRegex = RegExp(
    "^(S|D|W) ([0-9]{4}-[0-9]{2}-[0-9]{2}) ([0-9]{4}-[0-9]{2}-[0-9]{2})$"
  );

  function makeSheetName(kind: "S" | "D" | "W", from: Date, until: Date) {
    const fromS = DateUtils.toISODate(from);
    const untilS = DateUtils.toISODate(until);
    const res = `${kind} ${fromS} ${untilS}`;
    Logger.log(res);
    return res;
  }

  /** Create a new set of sheets and do their layout. */
  export function create(from: Date, until: Date): Item {
    const s = SpreadsheetApp.getActiveSpreadsheet();
    const scheduleSheetName = makeSheetName("S", from, until);
    const workSheetName = makeSheetName("W", from, until);
    const doodleSheetName = makeSheetName("D", from, until);
    const scheduleSheet = s.insertSheet(scheduleSheetName);
    const workSheet = s.insertSheet(workSheetName);
    const doodleSheet = s.insertSheet(doodleSheetName);
    SheetLayouter.setup(scheduleSheet, workSheet, doodleSheet, from, until);
    return { from, until, scheduleSheet, workSheet, doodleSheet };
  }

  /** Return the active item (aka the one that is currently selected). */
  export function getActiveItem(): Item | undefined {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getActiveSheet();
    const r = parseSheetName(sheet.getName());
    if (r === undefined) {
      return undefined;
    }
    const scheduleSheet = ss.getSheetByName(
      makeSheetName("S", r.from, r.until)
    );
    const doodleSheet = ss.getSheetByName(makeSheetName("D", r.from, r.until));
    const workSheet = ss.getSheetByName(makeSheetName("W", r.from, r.until));
    return {
      from: r.from,
      until: r.until,
      scheduleSheet,
      doodleSheet,
      workSheet,
    };
  }

  function parseSheetName(
    s: string
  ): { from: Date; until: Date; kind: "S" | "D" | "W" } | undefined {
    const m = nameRegex.exec(s);
    if (m !== null) {
      const from = new Date(m[2]);
      const until = new Date(m[3]);
      const kind = m[1] as "S" | "D" | "W";
      return { from, until, kind };
    }
    return undefined;
  }

  export function validateAndList(): Item[] {
    const items: Record<string, ItemInProgress> = {};
    const s = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = s.getSheets();
    sheets.forEach((s) => {
      const name = s.getName();
      const m = nameRegex.exec(name);
      if (m !== null) {
        const from = new Date(m[2]);
        const until = new Date(m[3]);
        const kind = m[1];
        const key = from.toString() + "-" + until.toString();
        if (!items[key]) {
          items[key] = { from, until };
        }
        const item = items[key];
        switch (kind) {
          case "S":
            item.scheduleSheet = s;
            break;
          case "D":
            item.doodleSheet = s;
            break;
          case "W":
            item.workSheet = s;
            break;
          default:
            throw Error("Can't happen because of regex above...");
        }
      }
    });

    const res: Item[] = [];
    for (const key of Object.keys(items)) {
      const item = items[key];
      if (item.scheduleSheet === undefined) {
        throw Error(`Das S sheet fuer ${item.from} - ${item.until} fehlt.`);
      }
      if (item.doodleSheet === undefined) {
        throw Error(`Das D sheet fuer ${item.from} - ${item.until} fehlt.`);
      }
      if (item.workSheet === undefined) {
        throw Error(`Das W sheet fuer ${item.from} - ${item.until} fehlt.`);
      }
      res.push({
        from: item.from,
        until: item.until,
        scheduleSheet: item.scheduleSheet!,
        doodleSheet: item.doodleSheet!,
        workSheet: item.workSheet!,
      });
    }

    return res;
  }
}

namespace ScheduleCount {
  function gridArrayLength(a: unknown[][]): number {
    return Math.max(a[0].length, a.length);
  }

  function gridArrayMap<A, B>(a: A[][], f: (elem: A) => B): B[][] {
    return a.map((a2) => a2.map(f));
  }

  /** The actual implementation of schedule count */
  function doScheduleCount(
    employees: string[],
    slot1DefaultMinutes: number,
    slot2DefaultMinutes: number,
    slot3DefaultMinutes: number,
    slots: string[]
  ): number[][] {
    const slotDefaults = [
      slot1DefaultMinutes,
      slot2DefaultMinutes,
      slot3DefaultMinutes,
    ];
    const nameToNdx: Record<string, number> = {};
    const result: number[] = employees.map((_) => 0);
    employees.forEach((name, ndx) => {
      nameToNdx[name] = ndx;
    });
    slots.forEach((slot, slotNdx) => {
      const slotItems = SlotParser.parse(slot);
      slotItems.forEach((i) => {
        let minutes: number = 0;
        switch (i.kind) {
          case "default":
            minutes = slotDefaults[slotNdx % 3];
            break;
          case "specified":
            minutes = i.duration;
            break;
        }
        const ndx = nameToNdx[i.name];
        if (ndx === undefined) {
          throw Error(`Unbekannter Mitarbeiter ${i.name}`);
        }
        result[ndx] += minutes;
      });
    });
    return [result];
  }

  function assertNumber(n: unknown, arg: string): number {
    if (typeof n === "number") {
      return n;
    } else {
      throw Error(`Expected ${arg} to be number.`);
    }
  }

  function assertGridOfStrings(a: unknown, arg: string): string[] {
    const result: string[] = [];
    if (a instanceof Array) {
      a.forEach((a2) =>
        a2.forEach((x: unknown) => {
          if (typeof x === "string") {
            result.push(x);
          } else {
            throw Error(`Expected ${arg} to be a range strings.`);
          }
        })
      );
    } else {
      throw Error(`Expected ${arg} to be a range strings.`);
    }
    return result;
  }

  function assertStrings(a: unknown, arg: string): string[] {
    if (a instanceof Array && a.every((s: unknown) => typeof s === "string")) {
      return a as string[];
    } else {
      throw Error(`Expected ${arg} to be a range strings.`);
    }
  }

  /** The exported entry point that does the type checking */
  export function main(
    employees: unknown,
    slot1DefaultMinutes: unknown,
    slot2DefaultMinutes: unknown,
    slot3DefaultMinutes: unknown,
    slots: unknown[]
  ): unknown {
    const slot1DefaultMinutesN = assertNumber(
      slot1DefaultMinutes,
      "slot1DefaultMinutes"
    );
    const slot2DefaultMinutesN = assertNumber(
      slot2DefaultMinutes,
      "slot2DefaultMinutes"
    );
    const slot3DefaultMinutesN = assertNumber(
      slot3DefaultMinutes,
      "slot3DefaultMinutes"
    );
    const employeesA = assertGridOfStrings(employees, "employees");
    const slotsA = assertStrings(slots, "slots");
    return doScheduleCount(
      employeesA,
      slot1DefaultMinutesN,
      slot2DefaultMinutesN,
      slot3DefaultMinutesN,
      slotsA
    );
  }
}

/**
 * Parse schedule slots.  Returns an array matching the dimensions of
 * the employee array where each element is the number of minutes worked
 * by the corresponding employee.
 * @param employees 1 dimensional array of employees
 * @param {number} slot1DefaultMinutes
 * @param {number} slot2DefaultMinutes
 * @param {number} slot3DefaultMinutes
 * @param {...string} slots
 * @customfunction
 */
function SCHEDULECOUNT(
  employees: unknown,
  slot1DefaultMinutes: unknown,
  slot2DefaultMinutes: unknown,
  slot3DefaultMinutes: unknown,
  ...slots: unknown[]
): unknown {
  return ScheduleCount.main(
    employees,
    slot1DefaultMinutes,
    slot2DefaultMinutes,
    slot3DefaultMinutes,
    slots
  );
}

namespace EditEventDecoder {
  /** A decoded version of the onEdit event:
   * OnEditInsert means the cell was previously empty
   * OnEditChange means the cell was not empty
   * OnEditDelete means the cell is now empty
   */
  export interface OnEditInsert {
    kind: "insert";
    value: any;
  }

  export interface OnEditChange {
    kind: "change";
    value: any;
    oldValue: any;
  }

  export interface OnEditClear {
    kind: "clear";
    oldValue: any;
    value: "";
  }

  export interface OnEditMassChange {
    kind: "mass-change";
  }

  export type OnEditEvent =
    | OnEditInsert
    | OnEditChange
    | OnEditClear
    | OnEditMassChange;

  /** Turn a google sheet onEdit Event into a typed event. */
  export function onEditEvent(
    event: GoogleAppsScript.Events.SheetsOnEdit
  ): OnEditEvent {
    if (event.oldValue === undefined && event.value === undefined) {
      return { kind: "mass-change" };
    } else if (event.oldValue === undefined && event.value !== undefined) {
      return { kind: "insert", value: event.value };
    } else if (
      event.oldValue !== undefined &&
      event.value.oldValue !== undefined
    ) {
      return { kind: "clear", oldValue: event.oldValue, value: "" };
    } else {
      return { kind: "change", oldValue: event.oldValue, value: event.value };
    }
  }
}

// Core idea here is that we detect additions and turn them into moves
// Note that comments on stackexchanges claim that onEdit events can
// get collapsed.  I couldn't observe this in testing, but it's better
// to not store critical functionality in the event handlers.
namespace EditHandler {
  function turnDuplicatesIntoMoves(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    slot: SheetLayouter.Slot
  ) {
    // We remove any duplicates of the words in
    // the current slots from all other slots on the same day.
    const entries = SheetLayouter.rangeOfEntriesOfDay(
      sheet,
      slot.date,
      "include-planner-exclude-doodle"
    );
    const cells = entries.range.getValues();
    const employeesInChangedCell = SlotParser.parse(
      cellAsString(cells[slot.row - entries.row][slot.column - entries.column])
    );
    const namesInChangedCell = employeesInChangedCell.map((e) => e.name);
    if (namesInChangedCell.length > 0) {
      let hasChange = false;
      for (let r = 0; r < cells.length; r++) {
        for (let c = 0; c < cells[r].length; c++) {
          if (
            (r !== slot.row - entries.row ||
              c !== slot.column - entries.column) &&
            cells[r][c] !== ""
          ) {
            // Todo hoist compilation of regex outside of the loop
            const newSlot = SlotParser.removeEmployees(
              cellAsString(cells[r][c]),
              namesInChangedCell
            );
            if (newSlot !== cells[r][c]) {
              cells[r][c] = newSlot;
              hasChange = true;
            }
          }
        }
      }
      if (hasChange) {
        entries.range.setValues(cells);
      }
    }
  }

  export function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
    const item = SheetsManager.getActiveItem();
    // Is an item currently active and was it's schedule sheet edited?
    const sheet = e.range.getSheet();
    if (
      item !== undefined &&
      sheet.getName() === item.scheduleSheet.getName()
    ) {
      const ev = EditEventDecoder.onEditEvent(e);

      switch (ev.kind) {
        case "mass-change":
          // After a mass change we just give up and hope the user will
          // recognize that they have to do manual work
          break;
        case "change":
        case "clear":
        case "insert":
          const slot = SheetLayouter.cellToSlot(
            sheet,
            "include-planner-exclude-doodle",
            {
              row: e.range.getRow(),
              column: e.range.getColumn(),
            }
          );
          if (slot !== undefined) {
            turnDuplicatesIntoMoves(sheet, slot);
          }
          break;
      }
    }
  }
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit): void {
  EditHandler.onEdit(e);
}

function promptForDate(
  ui: GoogleAppsScript.Base.Ui,
  prompt: string
): Date | undefined {
  const res = ui.prompt(
    prompt,
    "Bitte gebe ein Datum ein (YYYY-MM-TT zBsp 1981-07-14)",
    ui.ButtonSet.OK_CANCEL
  );
  const bt = res.getSelectedButton();
  if (bt === ui.Button.OK) {
    return new Date(res.getResponseText());
  } else {
    return undefined;
  }
}

function createNewItem(
  items: SheetsManager.Item[],
  from: Date,
  until: Date
): SheetsManager.Item {
  items.forEach((item) => {
    item.scheduleSheet.setTabColor(null);
    item.doodleSheet.setTabColor(null);
    item.workSheet.setTabColor(null);
  });
  const item = SheetsManager.create(from, until);
  item.doodleSheet.setTabColor("green");
  item.workSheet.setTabColor("green");
  item.scheduleSheet.setTabColor("green");
  item.scheduleSheet.activate();
  return item;
}

function menuCbNewSchedule() {
  const items = SheetsManager.validateAndList();
  const ui = SpreadsheetApp.getUi();
  const from = promptForDate(ui, "Von");
  if (from === undefined) {
    return;
  }
  const until = promptForDate(ui, "Bis");
  if (until === undefined) {
    return;
  }
  createNewItem(items, from, until);
}

function menuCbParseDoodle() {
  const ui = SpreadsheetApp.getUi();
  const items = SheetsManager.validateAndList();
  const entries = DoodleParser.parse();
  const dates = entries.map((e) => e.date);
  const from = dates.reduce(DateUtils.min);
  const until = dates.reduce(DateUtils.max);
  // this would not work well if nobody said they can work on the first or last day...
  let item = find(
    items,
    (i) => DateUtils.equal(i.from, from) && DateUtils.equal(i.until, until)
  );
  if (item === undefined) {
    const text =
      `Es gibt noch kein Tabellenblatt fuer den Zeitraum ${DateUtils.toISODate(
        from
      )} - ${DateUtils.toISODate(until)}\n` +
      `Soll ein neues Blatt erstellt werden?`;
    const res = ui.alert(text, ui.ButtonSet.OK_CANCEL);
    if (res !== ui.Button.OK) {
      return;
    }
    item = createNewItem(items, from, until);
  }
  SheetLayouter.replaceDoodle(item.scheduleSheet, entries);
}

function menuCbHideWD() {
  const items = SheetsManager.validateAndList();
  items.forEach((item) => {
    item.workSheet.hideSheet();
    item.doodleSheet.hideSheet();
  });
}

function menuCbShowWD() {
  const items = SheetsManager.validateAndList();
  items.forEach((item) => {
    item.workSheet.showSheet();
    item.doodleSheet.showSheet();
  });
}

function menuCbEmployeesChanged() {
  const items = SheetsManager.validateAndList();
  items.forEach((item) => {
    SheetLayouter.updateFormulas({
      scheduleSheet: item.scheduleSheet,
      doodleSheet: item.doodleSheet,
      workSheet: item.workSheet,
      fromDate: item.from,
      toDate: item.until,
    });
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu("BS")
    .addItem("Neu...", "menuCbNewSchedule")
    .addSeparator()
    .addItem("Doodle einlesen.", "menuCbParseDoodle")
    .addSeparator()
    .addItem("Verstecke W, D.", "menuCbHideWD")
    .addItem("Zeige W, D.", "menuCbShowWD")
    .addSeparator()
    .addItem("Mitarbeiter wurden geaendert.", "menuCbEmployeesChanged")
    .addToUi();
}
