namespace Entry {
  export interface IEntry {
    date: Date;
    employee: string;
    location: Locations.ILocation;
    shift: Shifts.IShift;
  }
}
