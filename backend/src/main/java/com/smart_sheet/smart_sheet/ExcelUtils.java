package com.smart_sheet.smart_sheet;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.InputStream;
import java.util.*;

public class ExcelUtils {

    public static List<DataRow> readExcel(InputStream is) throws Exception {

        List<DataRow> rows = new ArrayList<>();

        // Load the Excel file in memory and target the first sheet
        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        Iterator<Row> rowIterator = sheet.iterator();
        List<String> headers = new ArrayList<>();

        // Safely formats any cell type (numbers, dates, formulas) to a String
        DataFormatter formatter = new DataFormatter();

        // 1. Extract headers from the very first row (used as Velocity template keys)
        if (rowIterator.hasNext()) {
            Row headerRow = rowIterator.next();
            for (Cell cell : headerRow) {
                headers.add(formatter.formatCellValue(cell).trim());
            }
        }

        // 2. Iterate through the remaining rows to process the actual business data
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Map<String, String> rowData = new HashMap<>();

            // Map each cell's value to its corresponding header key
            for (int i = 0; i < headers.size(); i++) {
                Cell cell = row.getCell(i);

                // Failsafe read: ensures numeric/date values are parsed as plain text
                String value = formatter.formatCellValue(cell);
                rowData.put(headers.get(i), value);
            }

            // Encapsulate the mapped row into the domain model
            rows.add(new DataRow(rowData));
        }

        // Release resources to prevent memory leaks in the server
        workbook.close();

        return rows;
    }
}