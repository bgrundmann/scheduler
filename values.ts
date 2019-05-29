/** @OnlyCurrentDoc */
// Functions to deal with the values in cells as returned by getValue and getValues
namespace Values {
  export class ValueConversionError {
    public expected: string;
    public got: any;
    public row: number | undefined;
    public column: number | undefined;
    constructor(e: string, g: any) {
      this.expected = e;
      this.got = g;
    }
  }

  export function asDate(v: unknown): Date {
    if (v instanceof Date) {
      return v;
    }
    throw new ValueConversionError("Date", v);
  }

  export function asString(v: unknown): string {
    if (typeof v === "string") {
      return v;
    }
    throw new ValueConversionError("string", v);
  }

  export function asNumber(v: unknown): number {
    if (typeof v === "number") {
      return v;
    }
    throw new ValueConversionError("number", v);
  }

  // See https://stackoverflow.com/questions/17715841/how-to-read-the-correct-time-values-from-google-spreadsheet
  const spreadSheetEpoch = new Date("Dec 30, 1899 00:00:00");

  export function asInterval(v: unknown): Interval {
    if (v instanceof Date) {
      return Interval.ofMilliSeconds(v.getTime() - spreadSheetEpoch.getTime());
    }
    throw new ValueConversionError("Interval", v);
  }

  export function get<E>(values: unknown[][],
    row: number, col: number, conv: (x: unknown) => E): E {
    try {
      return conv(values[row][col]);
    } catch (e) {
      if (e instanceof ValueConversionError) {
        e.row = row;
        e.column = col;
      }
      throw e;
    }
  }
}
