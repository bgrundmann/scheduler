namespace Prelude {
  export function makeDictionary<T>(l: T[], keyfun: ((x: T) => string)): Record<string, T> {
    const d: Record<string, T> = {};
    l.forEach((e) => {
      d[keyfun(e)] = e;
    });
    return d;
  }

  interface IHasName {
    name: string;
  }

  export type Ordering = "lt" | "eq" | "gt";

  export type Comparator<T> = (a: T, b: T) => Ordering;

  export function compareBy<T, A>(get: (t: T) => A, comparator: Comparator<A>): Comparator<T> {
    return (t1: T, t2: T) => comparator(get(t1), get(t2));
  }

  export function lexiographic<T>(comparators: Array<Comparator<T>>): Comparator<T> {
    return (t1: T, t2: T) => {
      for (const compareField of comparators) {
        switch (compareField(t1, t2)) {
          case "lt": return "lt";
          case "gt": return "gt";
          case "eq": break;
        }
      }
      return "eq";
    };
  }

  export function stringCompare(a: string, b: string): Ordering {
    if (a < b) {
      return "lt";
    } else if (a > b) {
      return "gt";
    } else {
      return "eq";
    }
  }

  export function numberCompare(a: number, b: number): Ordering {
    if (a < b) {
      return "lt";
    } else if (a > b) {
      return "gt";
    } else {
      return "eq";
    }
  }

  export function findIndex<T>(length: number,
                               get: (index: number) => T,
                               predicate: (elem: T) => boolean, start?: number): number|undefined {
    for (let i = start || 0; i < length; i++) {
      if (predicate(get(i))) {
        return i;
      }
    }
    return undefined;
  }

  export function makeFindByName<T extends IHasName>(l: T[]): ((name: string) => T | undefined) {
    const d = makeDictionary(l, (e) => e.name );
    return ((name) => {
      return d[name];
    });
  }

  export function inRangeInclusive(n: number, low: number, upp: number): boolean {
    return low <= n && n <= upp;
  }

  export function forEachAsList<T>(forEach: (f: (x: T) => void) => void, optFilter?: (x: T) => boolean): T[] {
    const res: T[] = [];
    const filter = optFilter || ((x) => true );
    forEach((elem: T) => { if (filter(elem)) { res.push(elem); } });
    return res;
  }

  export function unwrap<T>(x: T|undefined|null): T {
    if (x !== undefined && x !== null) {
      return x;
    }
    throw new Error("Failed to unwrap");
  }

  export function flattenArray<T>(arrayOfArrays: T[][]): T[] {
    const empty: T[] = [];
    return empty.concat(...arrayOfArrays);
  }
}

interface Array<T> {
  intersperse(sep: T): T[];
}

Array.prototype.intersperse = function<T> (sep: T) {
  const result: T[] = [];
  this.forEach((x) => {
    if (result.length !== 0) {
      result.push(sep);
    }
    result.push(x);
  });
  return result;
};
