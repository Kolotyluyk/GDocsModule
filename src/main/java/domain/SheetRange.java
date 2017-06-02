package domain;

/**
 * Created by Сергій on 02.06.2017.
 */
public enum SheetRange {
    HEADER_RANGE("A3:T3"),
    DATA_RANGE("A4:Z1000"),
    COUNT_OF_DAY_RANGE("T1:T1"),
    EXCHANGE_RATE_RANGE("T3:T3");

    private final String name;

    public String getName() {
        return name;
    }

    SheetRange(String name) {
        this.name = name;
    }

}

