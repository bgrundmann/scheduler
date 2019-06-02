/** @OnlyCurrentDoc */
namespace Config {
  export const WEEKEND_COLOR = "#FFF2CC";
}

namespace Shifts {
  export const enum Kind { Morning, Afternoon, WholeDay }

  export function germanNameOfKind(k: Kind): string {
    return ["Vormittags", "Nachmittags", "Ganztags"][k];
  }

  export class Shift {
    constructor(public readonly start: Interval,
      public readonly stop: Interval,
      public readonly breakLength: Interval,
      public readonly kind: Kind) { }

    public toString() {
      return this.start.toHHMM() + "-" + this.stop.toHHMM();
    }
  }

  // Sadly prelude can not be used during the init of module Config
  // because of module evaluation order.  So this sadness is here to delay
  // the dependency on prelude after that module is properly initialized

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
        Logger.log("In the tricky case: %s - %s", start.getHours(), stop.getHours())
        return Kind.WholeDay;
      } else {
        return Kind.Afternoon;
      }
    }
  }

  export function create(start: Interval, stop: Interval, breakLength: Interval): Shift {
    const key = start.toString() + "-" + stop.toString() + "-" + breakLength.toString();
    const maybeRes = cache[key];
    if (maybeRes === undefined) {
      const res = new Shift(start, stop, breakLength, classify(start, stop));
      cache[key] = res;
      return res;
    } else {
      return maybeRes;
    }
  }
}
