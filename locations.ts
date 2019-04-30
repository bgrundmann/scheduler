namespace Locations {
  export interface ILocation {
    name: string;
    ndx: number;
  }
  const theList: ILocation[] = [
    { name : "Marktgasse"},
    { name : "Ammergasse" },
    { name : "Buero" },
    { name : "Extras" },
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
