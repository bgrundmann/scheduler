/** @OnlyCurrentDoc */
namespace Config {
  export const WEEKEND_COLOR = "#FFF2CC";
}

namespace Shifts {
  export const enum Kind { Morning, Afternoon, WholeDay }

  export interface Shift {
    name: string;
    start: number;
    stop: number;
    breakLength: number;
    kind: Kind;
  }
  export const firstHalf: Shift = {
    name : "Vormittags", start : 10, stop : 14, breakLength : 0,
    kind : Kind.Morning,
  };
  export const secondHalf: Shift = {
    name : "Nachmittags", start : 13, stop : 19, breakLength : 0,
    kind : Kind.Afternoon,
  };
  export const whole: Shift = {
    name : "Ganztags", start : 10, stop : 19, breakLength : 1,
    kind : Kind.WholeDay,
  };
  export const all = // in the order they appear on the data sheet
      [ Shifts.firstHalf, Shifts.secondHalf, Shifts.whole ];

  let byNameCache: ( (name: string) => Shift | undefined ) | undefined ;

  export function byName(name: string): Shift | undefined {
    if (!byNameCache) {
      byNameCache = Prelude.makeFindByName(Shifts.all);
    }
    return byNameCache(name);
  }
}
