/** @OnlyCurrentDoc */
namespace Config {
  export const WEEKEND_COLOR = "#FFF2CC";
}

namespace Shifts {
  export const enum Kind { Morning, Afternoon, WholeDay }

  export function germanNameOfKind(k: Kind): string {
    return ["Vormittags", "Nachmittags", "Ganztags"][k];
  }

  export interface Shift {
    start: Interval;
    stop: Interval;
    breakLength: Interval;
    kind: Kind;
  }

  // Sadly prelude can not be used during the init of module Shifts
  // because of module evaluation order.  So this sadness to cause
  // a dependency on prelude only after evaluation has happened

  let compareF: Prelude.Comparator<Shift> | undefined;

  export function compare(a: Shift, b: Shift) {
    if (compareF === undefined) {
      compareF = Prelude.lexiographic([
        Prelude.compareBy((s: Shift) => s.start, Interval.compare),
        Prelude.compareBy((s: Shift) => s.stop, Interval.compare),
        Prelude.compareBy((s: Shift) => s.breakLength, Interval.compare),
      ]);
    }
    return compareF(a, b);
  }

  function setupStandardShifts(): Shift[] {
    const firstHalf: Shift = {
      start: Interval.hhmm(10, 0), stop: Interval.hhmm(14, 0), breakLength: Interval.zero,
      kind: Kind.Morning,
    };
    const secondHalf: Shift = {
      start: Interval.hhmm(13, 0), stop: Interval.hhmm(19, 0), breakLength: Interval.zero,
      kind: Kind.Afternoon,
    };
    const whole: Shift = {
      start: Interval.hhmm(10, 0), stop: Interval.hhmm(19, 0), breakLength: Interval.hhmm(1, 0),
      kind: Kind.WholeDay,
    };
    const all = // in the order they appear on the data sheet
      [firstHalf, secondHalf, whole];
    return all;
  }

  const cache: Record<string, Shift> = {};

  function classify(start: Interval, stop: Interval): Kind {
    const startsEarly = start.getHours() <= 12;
    const stopsEarly = stop.getHours() <= 14;
    if (startsEarly) {
      if (stopsEarly) {
        return Kind.Morning;
      } else {
        return Kind.WholeDay;
      }
    } else {
      if (stopsEarly) {
        // This one is tricky...
        return Kind.WholeDay;
      } else {
        return Kind.Afternoon;
      }
    }
  }

  export function create(start: Interval, stop: Interval, breakLength: Interval) {
    const key = start.toString() + "-" + stop.toString() + "-" + breakLength.toString();
    const maybeRes = cache[key];
    if (maybeRes === undefined) {
      const res = { start, stop, breakLength, kind: classify(start, stop) };
      cache[key] = res;
      return res;
    } else {
      return maybeRes;
    }
  }
}
