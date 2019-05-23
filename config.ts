/** @OnlyCurrentDoc */
namespace Config {
  export const WEEKEND_COLOR = "#FFF2CC";
}

namespace Shifts {
  export const enum Kind { Morning, Afternoon, WholeDay }

  export interface Shift {
    name: string;
    start: Interval;
    stop: Interval;
    breakLength: Interval;
    kind: Kind;
  }
  function setupStandardShifts(): Shift[] {
    const firstHalf: Shift = {
      name: "Vormittags", start: Interval.hhmm(10, 0), stop: Interval.hhmm(14, 0), breakLength: Interval.zero,
      kind: Kind.Morning,
    };
    const secondHalf: Shift = {
      name: "Nachmittags", start: Interval.hhmm(13, 0), stop: Interval.hhmm(19, 0), breakLength: Interval.zero,
      kind: Kind.Afternoon,
    };
    const whole: Shift = {
      name: "Ganztags", start: Interval.hhmm(10, 0), stop: Interval.hhmm(19, 0), breakLength: Interval.hhmm(1, 0),
      kind: Kind.WholeDay,
    };
    const all = // in the order they appear on the data sheet
      [firstHalf, secondHalf, whole];
    return all;
  }

  let byNameCache: ((name: string) => Shift | undefined) | undefined;

  export function byName(name: string): Shift | undefined {
    if (!byNameCache) {
      byNameCache = Prelude.makeFindByName(setupStandardShifts());
    }
    return byNameCache(name);
  }
}
