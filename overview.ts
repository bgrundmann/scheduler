namespace OverviewSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Uebersicht");

  let ROWS_PER_ENTRY = 0; // inititalized during setup
  const COLUMNS_PER_ENTRY = 2;
  const FIRST_ENTRY_ROW = 2;
  const FIRST_ENTRY_COLUMN = 2;
  let mondayStartingFirstWeek: Date|undefined;
  let fromDate: Date|undefined;
  let toDate: Date|undefined;

  function entryPosition(date: Date) {
    const col = DateUtils.dayOfWeekStartingMonday(date);
    const days = DateUtils.daysBetween(Prelude.unwrap(mondayStartingFirstWeek), date);
    const row = Math.floor(days / 7);
    return { row : (row * ROWS_PER_ENTRY) + FIRST_ENTRY_ROW, col : (col * COLUMNS_PER_ENTRY) + FIRST_ENTRY_COLUMN };
  }

  const dowNames = ["Mo.", "Di.", "Mi.", "Do.", "Fr.", "Sa.", "So."];

  export function setup(fDate: Date, tDate: Date) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setFrozenRows(1);
    sheet.setHiddenGridlines(true);
    ROWS_PER_ENTRY = 1 + Locations.all().length;
    fromDate = fDate;
    toDate = tDate;
    mondayStartingFirstWeek = DateUtils.mondayStartingWeekContaining(fromDate);
    const weeks = Math.ceil(DateUtils.daysBetween(fromDate, toDate) / 7);
    sheet.getRange(1, FIRST_ENTRY_COLUMN, 1, 14).setValue(100);
    sheet.autoResizeColumns(FIRST_ENTRY_COLUMN, 14);
    for (let dow = 0; dow < 7; dow++) {
      sheet.getRange(FIRST_ENTRY_ROW - 1, FIRST_ENTRY_COLUMN + dow * COLUMNS_PER_ENTRY, 1, COLUMNS_PER_ENTRY)
      .mergeAcross()
      .setValue(dowNames[dow])
      .setBorder(true, true, true, true, false, false, "#000000",
          SpreadsheetApp.BorderStyle.SOLID);
    }
    const firstHalfRanges: GoogleAppsScript.Spreadsheet.Range[] = [];
    const secondHalfRanges: GoogleAppsScript.Spreadsheet.Range[] = [];
    DateUtils.forEachDay(fromDate, toDate, (date: Date) => {
      const pos = entryPosition(date);
      const box = sheet.getRange(pos.row, pos.col, ROWS_PER_ENTRY, COLUMNS_PER_ENTRY)
        .setBorder(true, true, true, true, false, false, "#000000",
          SpreadsheetApp.BorderStyle.SOLID);
      if (DateUtils.isWeekend(date)) {
        box.setBackground(Config.WEEKEND_COLOR);
      }
      // TODO: Fix null in setBorder problem
      sheet.getRange(pos.row + 1, pos.col, ROWS_PER_ENTRY - 1, COLUMNS_PER_ENTRY)
        .setBorder(null, null, null, null, true, true, "#dddddd",
          SpreadsheetApp.BorderStyle.SOLID);
      sheet.getRange(pos.row, pos.col, 1, COLUMNS_PER_ENTRY)
        .mergeAcross()
        .setNumberFormat("d")
        .setHorizontalAlignment("center")
        .setValue(date);
      const formulas = Locations.all().map((loc) => {
        return ScheduleSheet.formulasEmployeeCount(date, loc);
      });
      sheet.getRange(pos.row + 1, pos.col, Locations.all().length, COLUMNS_PER_ENTRY)
        .setFormulas(formulas)
        .setNumberFormat("0")
        .setHorizontalAlignment("center");
      // the -1 is a hack to exclude the extras location
      // TODO: make location numbers of configurable
      firstHalfRanges.push(sheet.getRange(pos.row + 1, pos.col, Locations.all().length - 1, 1));
      secondHalfRanges.push(sheet.getRange(pos.row + 1, pos.col + 1, Locations.all().length - 1, 1));
    });
    const firstHalfVeryBadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setFontColor("red")
      .setBold(true)
      .setRanges(firstHalfRanges)
      .build();
    const secondHalfVeryBadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(0)
      .setFontColor("red")
      .setBold(true)
      .setRanges(secondHalfRanges)
      .build();
    const firstHalfBadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberEqualTo(1)
      .setFontColor("orange")
      .setBold(true)
      .setRanges(firstHalfRanges)
      .build();
    const secondHalfBadRule = SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 2)
      .setFontColor("orange")
      .setBold(true)
      .setRanges(secondHalfRanges)
      .build();
    sheet.setConditionalFormatRules([firstHalfVeryBadRule, secondHalfVeryBadRule, firstHalfBadRule, secondHalfBadRule]);
    const lastRow = entryPosition(toDate).row;
    for (let row = FIRST_ENTRY_ROW + 1; row <= lastRow + 1; row += ROWS_PER_ENTRY) {
      const locnames = Locations.all().map((l) => [l.name]);
      sheet.getRange(row, FIRST_ENTRY_COLUMN - 1, locnames.length, 1)
        .setValues(locnames)
        .setHorizontalAlignment("right");
    }
  }
}

function testOverview() {
  Logger.clear();
  OverviewSheet.setup(new Date("2019-04-11"), new Date("2019-05-22"));
}
