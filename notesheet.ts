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

    export function add(notes: Note[]): void {
        const rows = notes.map((n) => [n.date, n.index, n.text]);
        sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
        const data = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (data) {
            data.sort([{ column: 1, ascending: true }, { column: 2, ascending: true }]);
        }
    }

    function deleteRange(d1: Date, d2: Date): void {
        const range = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (!range) { return; }
        const data = range.getValues();
        const len = data.length;
        function get(n: number) {
            return Values.get(data, n, 0, Values.asDate);
        }
        const firstRow = Prelude.findIndex(len, get, (d) => DateUtils.inRangeInclusive(d, d1, d2));
        if (firstRow !== undefined) {
            const firstOtherRow =
                Prelude.findIndex(data.length, get, (d) => DateUtils.compare(d, d2) === "gt", firstRow + 1);
            const num = (firstOtherRow === undefined) ? len - firstRow : firstOtherRow - firstRow;
            // +1 because of 1 based-ness, +1 because of header
            sheet.deleteRows(firstRow + 1 + 1, num);
        }
    }

    /** Replace all notes for the dates in the given range (inclusive). */
    export function replaceRange(d1: Date, d2: Date, notes: Note[]): void {
        deleteRange(d1, d2);
        add(notes);
    }

    export function findRowMatching(date: Date, index: number): number|undefined {
        const range = SheetUtils.getDataRangeWithoutHeader(sheet);
        if (!range) { return undefined; }
        const data = range.getValues();
        const len = data.length;
        for (let rowNumber = 0; rowNumber < len; rowNumber++) {
            const row = data[rowNumber];
            const rowDate = Values.asDate(row[0]);
            const rowIndex = Values.asNumber(row[1]);
            switch (DateUtils.compare(date, rowDate)) {
                case "eq":
                    if (rowIndex === index) {
                         // + 1 cause arrays are 0 based and + 1 for the header
                        return rowNumber + 1 + 1;
                    }
                    break;
                case "lt":
                    return undefined;
            }
        }
        return undefined;
    }

    /** Delete the note with matching date and index (if any) */
    export function deleteMatching(date: Date, index: number): void {
        const row = findRowMatching(date, index);
        if (row !== undefined) {
            sheet.deleteRow(row);
        }
    }

    /** Add a new note or replace an existing one on same date and index */
    export function addOrReplace(note: Note): void {
        const row = findRowMatching(note.date, note.index);
        if (row !== undefined) {
            sheet.getRange(row, 3).setValue(note.text);
        } else {
            add([note]);
        }
    }
}
