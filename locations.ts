namespace Locations {
  export interface ILocation {
    name: string;
    ndx: number;
  }
  const theList: ILocation[] = [
    { name : "Ammergasse" },
    { name : "Buero" },
    { name : "Marktgasse"},
    { name : "Xtras" },
  ].map ((l, ndx) => ({ ...l, ndx }) );

  let byNameCache: ((name: string) => ILocation|undefined)|undefined;

  export function all(): ILocation[] { return theList; }

  export function byName(name: string): ILocation | undefined {
    if (!byNameCache) {
      byNameCache = Prelude.makeFindByName(theList);
    }
    return byNameCache(name);
  }
}
