namespace EmployeeSheet {
  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getSheetByName("Mitarbeiter");

  export interface IEmployee {
    employee: string;
    alias: string;
  }

  export function forEach(f: (e: IEmployee) => void): void {
    const dataRange = sheet.getDataRange().getValues();
    const rows = dataRange.length;
    for (let row = 2; row < rows; row++) {
      f({ employee: Values.get(dataRange, row, 0, Values.asString),
        alias: Values.get(dataRange, row, 1, Values.asString),
      });
    }
  }

  export function makeByAliasAndHandle(list: IEmployee[]) {
    const dict: Record<string, IEmployee> = {};
    list.forEach((e) => {
      dict[e.employee] = e;
      if (e.alias) {
        dict[e.alias] = e;
      }
    });
    return dict;
  }

  let allCached: IEmployee[] | undefined;
  let byAliasAndHandleCached: Record<string, IEmployee>|undefined;

  export function all(): IEmployee[] {
    if (!allCached) {
      allCached = Prelude.forEachAsList(forEach);
    }
    return allCached;
  }

  export function byAliasAndHandle(): Record<string, IEmployee> {
    if (!byAliasAndHandleCached) {
      byAliasAndHandleCached = makeByAliasAndHandle(all());
    }
    return byAliasAndHandleCached;
  }
}
