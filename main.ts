/** @OnlyCurrentDoc */
// TODO:
//   - Write something to print
//   - Write something to calendarize
//   - Add config page with list of employees and their powers
//   - Write alternative views
//     - A view that shows several weeks at once and only the number of employees per office
//   - Figure out how to deploy this properly
//     Core idea: deploy as sheet with bound script (but develop using typescript and clasp)
//     Deploy by making named copies of the sheet.
//     Debug by making backups and exporting that sheet
//   - Write something to make shifts that are not in the doodle bold
//   - Split Daten into Focus and History
//   - Use regular coloring instead of conditional formatting on the schedule sheet for weekends
namespace Main {
  export function saveEntriesFromScheduleToData() {
    const range = ScheduleSheet.dateRange();
    const existing: Entry.IEntry[] = [];
    ScheduleSheet.forEachEntry((e) => { existing.push(e); });
    DataSheet.replaceRange(range.from, range.until, existing);
  }
  export function recompute() {
    saveEntriesFromScheduleToData();
  }
  export function onOpenCallback() {
    Logger.clear();
    // Make sure whatever was left on the schedule page last time is saved into the data sheet
    saveEntriesFromScheduleToData();
    // Now that is done reset the data sheet
    // const d1 = new Date("2019-04-11");
    // const d2 = new Date("2019-05-22");
    // ScheduleSheet.setup(d1, d2);
    // OverviewSheet.setup(d1, d2);
  }
  export function changeDates() {
    const fromDate = SheetUtils.askForDate("Von");
    if (!fromDate) { return; }
    const untilDate = SheetUtils.askForDate("Bis");
    if (!untilDate) { return; }
    ScheduleSheet.setup(fromDate, untilDate);
    OverviewSheet.setup(fromDate, untilDate);
  }
  export function parseDoodle() {
    DoodleParser.parse();
  }
  export function employeesFromDoodleToSchedule() {
    const range = ScheduleSheet.dateRange();
    const whoAndWhere = ScheduleSheet.employeesAndLocations();
    const entriesToPlace = Prelude.forEachAsList(PollSheet.forEachUnique, (p) => {
      return whoAndWhere[p.employee] && range.from <= p.date && p.date <= range.until;
    });
    const entries = entriesToPlace.map((ps) => {
      return {
        employee : ps.employee,
        date : ps.date,
        // TODO: Handle error properly
        location : Prelude.unwrap(whoAndWhere[ps.employee].location),
        shift : ps.shift,
      };
    });
    DataSheet.add(entries);
    ScheduleSheet.setup(range.from, range.until);
  }
  export function backup() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    SheetUtils.deleteSheetByNameIfExists("Backup");
    const sheet = ss.getSheetByName("Daten").copyTo(ss);
    sheet.setName("Backup");
  }
  export function restore() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const backupSheet = ss.getSheetByName("Backup");
    if (!backupSheet) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Kein Backup vorhanden!");
    }
    const daten = ss.getSheetByName("Daten");
    daten.clear();
    backupSheet.getDataRange().copyTo(daten.getRange(1, 1));
    const r = ScheduleSheet.dateRange();
    ScheduleSheet.setup(r.from, r.until);
  }
  export function onEditCallback(e: GoogleAppsScript.Events.SheetsOnEdit) {
    Logger.log("onEditCallback %s", e);
    switch (e.range.getSheet().getName()) {
      case ScheduleSheet.NAME:
        ScheduleSheet.onEditCallback(e);
        break;
      default:
        Logger.log("No callback for this sheet");
    }
  }
}

function zeitraumAendernCallback() {
  Main.changeDates();
}

function neuBerechnenCallback() {
  Main.recompute();
}

function doodleEinlesenCallback() {
  if (SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Umfrage") === null) {
    SpreadsheetApp.getUi()
      .alert("Zuerst den doodle in dieses spreadsheet importieren" +
        " (Datei -> Importieren (WICHTIG: Neues Tabellenblatt einfuegen!))");
    return;
  }
  Main.parseDoodle();
}

function mitarbeiterUebertragenCallback() {
  Main.employeesFromDoodleToSchedule();
}

function sicherheitskopieErstellenCallback() {
  Main.backup();
}

function sicherheitskopieWiederherstellenCallback() {
  Main.restore();
}

function onOpen() {
  SpreadsheetApp.getUi()
  .createMenu("BS")
  .addItem("Neu berechnen", "neuBerechnenCallback")
  .addItem("Mitarbeiter Doodle -> Schedule", "mitarbeiterUebertragenCallback")
  .addSeparator()
  .addItem("Zeitraum aendern", "zeitraumAendernCallback")
  .addItem("Doodle einlesen", "doodleEinlesenCallback")
  .addSeparator()
  .addItem("Sicherheitskopie erstellen", "sicherheitskopieErstellenCallback")
  .addItem("Sicherheitskopie wiederherstellen!", "sicherheitskopieWiederherstellenCallback")
  .addToUi();
  Main.onOpenCallback();
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  Main.onEditCallback(e);
}

function initialSetup() {
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet("Schedule");
  spreadsheet.insertSheet("Uebersicht");
  spreadsheet.insertSheet("Daten");
  spreadsheet.insertSheet("Mitarbeiter");
}
