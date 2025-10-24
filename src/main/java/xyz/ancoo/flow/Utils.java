package xyz.ancoo.flow;

import java.math.BigDecimal;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

public class Utils {

    // 工具方法：将 Cell 转为字符串（处理各种类型）
    public static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 如果是日期格式
                    return cell.getDateCellValue().toString();
                } else {
                    // 普通数字，避免科学计数法，保留原样
                    double value = cell.getNumericCellValue();
                    // 如果是整数，转为整数字符串（如 100.0 → "100"）
                    if (value == (long) value) {
                        return String.valueOf((long) value);
                    } else {
                        return String.valueOf(value);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // 可选择计算公式结果或返回公式字符串
                try {
                    return getCellValueAsString(cell); // 递归获取计算后的值（不推荐，可能循环）
                    // 更安全做法：用 FormulaEvaluator
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default:
                return "";
        }
    }

    public static String parseDay(String value) {
        Pattern p = Pattern.compile("日期：(\\d{4})年(\\d{1,2})月(\\d{1,2})日");
        Matcher m = p.matcher(value);
        if (m.find()) {
            String year = m.group(1);
            String month = String.format("%02d", Integer.parseInt(m.group(2)));
            String day = String.format("%02d", Integer.parseInt(m.group(3)));
            return year + month + day;
        }
        throw new RuntimeException("日期格式错误");
    }

    public static String parseStartTime(String value) {
        Pattern p = Pattern.compile(".*时间：(\\d{1,2}):(\\d{2})");
        Matcher m = p.matcher(value);
        if (m.find()) {
            String hour = String.format("%02d", Integer.parseInt(m.group(1)));
            String minute = m.group(2); // 已是两位
            return hour + "." +  minute;
        }
        throw new RuntimeException("时间格式错误");
    }

    public static String parseEndTime(String value) {
        return parseStartTime(value);
    }

    public static BigDecimal parseStartFlow(String value) {
        return new BigDecimal(value.replace("m", "").trim());
    }

    public static BigDecimal parseEndFlow(String value) {
        Pattern p = Pattern.compile(".*水位：([+-]?\\d+\\.\\d+)m");
        Matcher m = p.matcher(value);
        if (m.find()) {
            return new BigDecimal(m.group(1));
        }
        throw new RuntimeException("水位格式错误");
    }
}
