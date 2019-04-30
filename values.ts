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

  export function asDate(v: any): Date {
    if (v instanceof Date) {
      return v;
    }
    throw new ValueConversionError("Date", v);
  }

  export function asString(v: any): string {
    if (typeof v === "string") {
      return v;
    }
    throw new ValueConversionError("string", v);
  }

  export function asNumber(v: any): number {
    if (typeof v === "number") {
      return v;
    }
    throw new ValueConversionError("number", v);
  }

  export function get<E>(values: any[][],
                         row: number, col: number, conv: (x: any) => E): E {
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
