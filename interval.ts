const MILLIS_PER_SECOND = 1000;
const MILLIS_PER_MINUTE = MILLIS_PER_SECOND * 60;
const MILLIS_PER_HOUR = MILLIS_PER_MINUTE * 60;

class Interval {
    /** Compute a - b (result can be negative). */
    public static diff(a: Interval, b: Interval): Interval {
        return new Interval(a.ms - b.ms);
    }

    /** Comparison */
    public static compare(a: Interval, b: Interval): Prelude.Ordering {
        return Prelude.numberCompare(a.ms, b.ms);
    }

    /** Create a new interval */
    public static hhmm(hours: number, minutes: number): Interval {
        return new Interval(hours * MILLIS_PER_HOUR + minutes * MILLIS_PER_MINUTE);
    }

    protected constructor(private readonly ms: number) { }

    public getHours(): number {
        return Math.floor(this.ms / MILLIS_PER_HOUR);
    }

    public getMinutes(): number {
        return Math.floor((this.ms % MILLIS_PER_HOUR) / MILLIS_PER_MINUTE);
    }

    public toString(): string {
        return this.ms.toString();
    }
}

function testInterval() {
    function eq(a: unknown, b: unknown): void {
        if (a !== b) {
            throw Error(`Expected ${a} to be equal to ${b}`);
        }
    }
    const a = Interval.hhmm(8, 0);
    const b = Interval.hhmm(12, 0);
    const c = Interval.hhmm(9, 45);

    eq(a.getHours(), 8);
    eq(a.getMinutes(), 0);
    eq(b.getHours(), 12);
    eq(b.getMinutes(), 0);
    eq(c.getHours(), 9);
    eq(c.getMinutes(), 45);

    eq(Interval.diff(b, a).getHours(), 4);
    eq(Interval.diff(b, a).getMinutes(), 0);
    eq(Interval.diff(c, a).getHours(), 1);
    eq(Interval.diff(c, a).getMinutes(), 45);
    eq(Interval.diff(b, c).getMinutes(), 15);
    eq(Interval.compare(a, b), "lt");
    eq(Interval.compare(b, a), "gt");
    eq(Interval.compare(a, Interval.hhmm(8, 0)), "eq");
}
