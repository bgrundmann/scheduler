/** @OnlyCurrentDoc */
// TODO:
//   - Write something to print
//   - Write something to calendarize
//   - Add config page with list of employees and their powers
//   - Figure out how to deploy this properly
//     Core idea: deploy as sheet with bound script (but develop using typescript and clasp)
//     Deploy by making named copies of the sheet.
//     Debug by making backups and exporting that sheet
//   - Write something to make shifts that are not in the doodle bold
//   - Split Daten into Focus and History
//   - Add column to show special dates (maybe by subscribing to a calendar?)
//   - Make cells that are completely empty red
//   - more clever doodle -> schedule rules
//   - doodle box?
//   - Samstags sind nur 6 stunden
//     Anfang: 9:45, ende: 1900.  Samstags ende: 1600.  Sonntags: 13:00 - 18:00 (aber mal 1.5)
//   - Wie sollte man schulungen verrechnen?
//
namespace Main {
  export function onOpenCallback() {
    Logger.clear();
    ScheduleSheet.syncScheduleToData();
  }
  export function changeDates() {
    // make sure everything is saved first.
    ScheduleSheet.syncScheduleToData();
    const fromDate = SheetUtils.askForDate("Von");
    if (!fromDate) {
      return;
    }
    const untilDate = SheetUtils.askForDate("Bis");
    if (!untilDate) {
      return;
    }
    ScheduleSheet.setup(fromDate, untilDate);
  }
  export function parseDoodle() {
    DoodleParser.parse();
  }
  export function employeesFromDoodleToSchedule() {
    ScheduleSheet.syncScheduleToData();
    const range = ScheduleSheet.dateRange();
    const whoAndWhere = ScheduleSheet.employeesAndLocations();
    const entriesToPlace = Prelude.forEachAsList(
      PollSheet.forEachUnique,
      (p) => {
        return (
          whoAndWhere[p.employee] &&
          DateUtils.inRangeInclusive(p.date, range.from, range.until)
        );
      }
    );
    const entries: Entry.IEntry[] = entriesToPlace.map((ps) => {
      return {
        employees: [ps.employee],
        date: ps.date,
        location: Prelude.unwrap(whoAndWhere[ps.employee].location),
        shift: ps.shift,
      };
    });
    DataSheet.add(entries);
    ScheduleSheet.setup(range.from, range.until);
  }
  export function backup() {
    ScheduleSheet.syncScheduleToData();
    SheetUtils.backupSheet("Daten", "Backup");
    SheetUtils.backupSheet("Notizen", "Backup-Notizen");
  }
  export function restore() {
    SheetUtils.restoreSheet("Backup-Notizen", "Notizen");
    if (!SheetUtils.restoreSheet("Backup", "Daten")) {
      const ui = SpreadsheetApp.getUi();
      ui.alert("Kein Backup vorhanden!");
    }
    const r = ScheduleSheet.dateRange();
    ScheduleSheet.setup(r.from, r.until);
  }
  export function onEditCallback(e: GoogleAppsScript.Events.SheetsOnEdit) {
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

function doodleEinlesenCallback() {
  if (
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Umfrage") === null
  ) {
    SpreadsheetApp.getUi().alert(
      "Zuerst den doodle in dieses spreadsheet importieren" +
        " (Datei -> Importieren (WICHTIG: Neues Tabellenblatt einfuegen!))"
    );
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
    .addItem("Mitarbeiter Doodle -> Schedule", "mitarbeiterUebertragenCallback")
    .addSeparator()
    .addItem("Zeitraum aendern", "zeitraumAendernCallback")
    .addItem("Doodle einlesen", "doodleEinlesenCallback")
    .addSeparator()
    .addItem("Sicherheitskopie erstellen", "sicherheitskopieErstellenCallback")
    .addItem(
      "Sicherheitskopie wiederherstellen!",
      "sicherheitskopieWiederherstellenCallback"
    )
    .addToUi();
  Main.onOpenCallback();
}

function onEdit(e: GoogleAppsScript.Events.SheetsOnEdit) {
  Logger.log("onEdit start");
  Main.onEditCallback(e);
  Logger.log("onEdit stop");
}

function initialSetup1() {
  const spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.insertSheet("Schedule");
  spreadsheet.insertSheet("Notizen");
  spreadsheet.insertSheet("Daten");
  spreadsheet.insertSheet("Mitarbeiter");
  DataSheet.initialSetup();
}

function initialSetup2() {
  ScheduleSheet.setup(new Date("2019-05-23"), new Date("2019-06-26"));
}
