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

  export function makeFindByName<T extends IHasName>(l: T[]): ((name: string) => T | undefined) {
    const d = makeDictionary(l, (e) => e.name );
    return ((name) => {
      return d[name];
    });
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
