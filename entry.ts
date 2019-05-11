namespace Entry {
  /** A place is the combination of a date and location.
   * For example in the Office on the 20th of March 2019
   */
  export interface Place {
    readonly date: Date;
    readonly location: Locations.ILocation;
  }

  export function samePlace(p1: Place, p2: Place): boolean {
    return DateUtils.equal(p1.date, p2.date) && p1.location.name === p2.location.name;
  }

  /** A slot is the combination of a Place and a shift.
   */
  export interface Slot extends Place {
    readonly shift: Shifts.IShift;
  }

  export function sameSlot(s1: Slot, s2: Slot): boolean {
    return samePlace(s1, s2) && s1.shift.name === s2.shift.name;
  }

  /** A single entry in the schedule corresponds to all
   * employees working a particular slot.
   * single person working a particular slot.
   */
  export interface IEntry extends Slot {
    readonly employees: string[];
  }
}
