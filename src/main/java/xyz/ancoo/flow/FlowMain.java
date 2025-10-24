package xyz.ancoo.flow;

import static xyz.ancoo.flow.Utils.getCellValueAsString;
import java.io.File;
import java.io.FileInputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class FlowMain {

    public static void main(String[] args) throws Exception {
        args = new String[] {
                "C:\\Users\\liuchangqing\\Downloads\\tmp\\81600450.zwl",
                "C:\\Users\\liuchangqing\\Downloads\\tmp"
        };

        if (args.length != 2) {
            System.out.println("Usage: java -jar flow.jar D://xls//81600450.zwl D://xls//");
            return;
        }

        Map<String, ZWLData> datas = parseZWL(args[0]);
        if (datas == null || datas.isEmpty()) {
            System.out.println("zwl file not exists");
            return;
        }
        System.out.println("zwl data count: " + datas.size());

        File xlsDir = new File(args[1]);
        if (!xlsDir.exists() || !xlsDir.isDirectory()) {
            System.out.println("xls directory not exists or content is empty");
            return;
        }

        File[] files = xlsDir.listFiles();
        if (files == null || files.length < 1) {
            System.out.println("xls directory is empty");
            return;
        }

        for (File file : files) {
            handleXls(datas, file);
        }
    }

    private static void handleXls(Map<String, ZWLData> datas, File file) throws Exception {
        if (!file.getName().endsWith(".xls")) {
            return;
        }
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new HSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            String id = "";

            // 日期：2025年1月1日
            int rowIndex = 2;
            int colIndex = 0;
            String value = getCellValue(file.getName(), sheet, rowIndex, colIndex);
            id += Utils.parseDay(value);

            // 开始时间：13:04
            rowIndex = 4;
            colIndex = 0;
            value = getCellValue(file.getName(), sheet, rowIndex, colIndex);
            ZWLData zwlStartData = datas.get(id + Utils.parseStartTime(value));

            // -2.21m 开始水位
            rowIndex = 29;
            colIndex = 3;
            value = getCellValue(file.getName(), sheet, rowIndex, colIndex);
            BigDecimal startFlow = Utils.parseStartFlow(value);

            // 结束时间：13:11
            rowIndex = 4;
            colIndex = 4;
            value = getCellValue(file.getName(), sheet, rowIndex, colIndex);
            ZWLData zwlEndData = datas.get(id + Utils.parseEndTime(value));

            rowIndex = 29;
            colIndex = 4;
            value = getCellValue(file.getName(), sheet, rowIndex, colIndex);
            BigDecimal endFlow = Utils.parseEndFlow(value);

            System.out.printf("%s-> 开始水位: %s, 结束水位: %s, 开始水位（ZWL）: %s, 结束水位（ZWL）: %s, dif: %s %s \n",
                    file.getName(),
                    startFlow,
                    endFlow,
                    zwlStartData == null ? "-" : zwlStartData.value(),
                    zwlEndData == null ? "-" : zwlEndData.value(),
                    zwlStartData == null ? "-" : zwlStartData.value().subtract(startFlow),
                    zwlEndData == null ? "-" : zwlEndData.value().subtract(endFlow));
        }
    }

    public static String getCellValue(String file, Sheet sheet, int rowIndex, int colIndex) {
        String value = getCellValue(sheet, rowIndex, colIndex);
        if (value == null || value.isEmpty()) {
            System.out.printf("File format error. File: %s, row: %d, col: %d \n",
                    file, rowIndex, colIndex);
            System.exit(0);
        }
        return value;
    }


    public static String getCellValue(Sheet sheet, int rowIndex, int colIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            return null;
        }

        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            System.out.println("单元格为空");
            return null;
        }

        return getCellValueAsString(cell);
    }

    private static Map<String, ZWLData> parseZWL(String file) throws Exception {
        Path path = Paths.get(file);
        if (!Files.exists(path)) {
            System.out.println("zwl file not exists");
            return null;
        }

        Map<String, ZWLData> datas = new HashMap<>();
        ZWLData lastData = null;
        for (String line : Files.readAllLines(path)) {
            ZWLData data = ZWLData.parse(line);
            if (data == null) {
                continue;
            }
            datas.put(data.id(), data);

            if (lastData == null) {
                lastData = data;
                continue;
            }
            // 内插
            if (!lastData.day().equals(data.day())) {
                lastData = data;
                continue;
            }
            if (lastData.minute() == data.minute()) {
                lastData = data;
                continue;
            }
            BigDecimal step = data.value().subtract(lastData.value())
                    .divide(BigDecimal.valueOf(data.minute() - lastData.minute()),
                            2, RoundingMode.HALF_UP);
            for (int i = lastData.minute() + 1; i < data.minute(); i++) {
                BigDecimal newValue = lastData.value().add(step.multiply(new BigDecimal(i - lastData.minute())));
                ZWLData newData = new ZWLData(data.day(), i, newValue);
                datas.put(newData.id(), newData);
            }
            lastData = data;
        }
        return datas;
    }
}