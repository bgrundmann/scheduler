/** @OnlyCurrentDoc */
namespace DateUtils {
  export function copy(d: Date): Date { return new Date(d.getTime()); }
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
  export function inRangeInclusive<T>(d: T, low: T, upp: T): boolean {
    if (typeof d === "number") {
      return low <= d && d <= upp;
    } else if (d instanceof Date && low instanceof Date && upp instanceof Date) {
      return low.getTime() <= d.getTime() && d.getTime() <= upp.getTime();
    }
    // never happens just to make compiler happy
    return false;
  }
  export function equal(d1: Date, d2: Date): boolean {
    return d1.getTime() === d2.getTime();
  }
  export function forEachDay(lower: Date, upper: Date, f: (d: Date) => void): void {
    let d = lower;
    while (d <= upper) {
      f(d);
      d = nextDay(d);
    }
  }
  export function mondayStartingWeekContaining(d: Date): Date {
    const res = copy(d);
    while (res.getDay() !== 1) { res.setDate(res.getDate() - 1); }
    return res;
  }
  export function daysBetween(d1: Date, d2: Date): number {
    return Math.round((d2.getTime() - d1.getTime()) / (1000 * 60 * 60 * 24));
  }
  const isoRegex = /\d{4}-\d{2}-\d{2}/;

  /*** Parse a date in extended iso format */
  export function parseISODate(s: string): Date|undefined {
    const r = isoRegex.exec(s);
    if (r !== null) {
      return new Date(s);
    }
    return undefined;
  }
}
