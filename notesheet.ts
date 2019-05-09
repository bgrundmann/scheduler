namespace NoteSheet {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getSheetByName("Notizen");

    export interface Note {
        date: Date;
        index: number;
        text: string;
    }

    /** Call f foreach note in the given range (inclusive) */
    export function forEachEntryInRange(d1: Date, d2: Date, f: (note: Note) => void): void {
        const data = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (data) {
            data.getValues().forEach((row) => {
                const date = Values.asDate(row[0]);
                if (DateUtils.inRangeInclusive(date, d1, d2)) {
                    const index = Values.asNumber(row[1]);
                    const text = Values.asString(row[2]);
                    f({date, index, text});
                }
            });
        }
    }

    export function add(notes: Note[]) {
        const rows = notes.map((n) => [n.date, n.index, n.text]);
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
        const data = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (data) {
            data.sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
        }
    }

    export function findRowMatching(date: Date, index: number): number|undefined {
        const range = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (!range) { return undefined; }
        const data = range.getValues();
        const len = data.length;
        for (let rowNumber = 0; rowNumber < len; rowNumber++) {
            const row = data[rowNumber];
            const candidateDate = Values.asDate(row[0]);
            const candidateIndex = Values.asNumber(row[1]);
            if (DateUtils.equal(date, candidateDate) && candidateIndex === index) {
                /// + 1 cause arrays are 0 based and + 1 for the header
                return rowNumber + 1 + 1;
            } else if (date.getTime() > candidateDate.getTime()) {
                return undefined;
            }
        }
        return undefined;
    }

    /** Add a new note or replace an existing one on same date and index */
    export function addOrReplace(note: Note) {
        const row = findRowMatching(note.date, note.index);
        if (row !== undefined) {
            sheet.getRange(row, 3).setValue(note.text);
        } else {
            add([note]);
        }
    }
}
