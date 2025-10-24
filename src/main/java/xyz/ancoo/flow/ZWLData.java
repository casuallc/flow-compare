package xyz.ancoo.flow;

import java.math.BigDecimal;

public class ZWLData {

    private String day;

    private int minute;

    private BigDecimal value;

    public ZWLData() {
    }

    public ZWLData(String day, int minute, BigDecimal value) {
        this.day = day;
        this.minute = minute;
        this.value = value;
    }

    public static ZWLData parse(String line) {
        if (line == null || line.isEmpty()) {
            return null;
        }
        ZWLData data = new ZWLData();
        String[] parts = line.split("\\s+");
        String[] dates = parts[0].split("\\.");
        data.day = dates[0].trim();
        data.minute = Integer.parseInt(dates[1].trim());

        data.value = new BigDecimal(parts[1].trim());
        return data;
    }

    public String id() {
        return day + "." + (minute < 10 ? "0" + minute : minute);
    }

    public String day() {
        return day;
    }

    public int minute() {
        return minute;
    }

    public BigDecimal value() {
        return value;
    }
}
