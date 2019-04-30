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
}
