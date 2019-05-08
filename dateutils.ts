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
  export function inRangeInclusive<T extends number|Date>(d: T, low: T, upp: T): boolean {
    return low <= d && d <= upp;
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
  const monthRegex = /\d{4}-\d{2}/;
  const dateRangeRegex = /\d{4}-\d{2}-\d{2} - \d{4}-\d{2}-\d{2}/;

  export function parseHumanDateRangeInput(d: string): { from: Date, until: Date} | null {
    const r = monthRegex.exec(d);
    if (r !== null) {
      const d1 = new Date(d + "-01");
      const d2 = new Date(d1.getFullYear(), d1.getMonth() + 2, 0);
      return { from : d1, until : d2 };
    }
    return null;
  }
}
