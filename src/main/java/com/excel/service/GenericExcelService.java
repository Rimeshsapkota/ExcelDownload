package com.excel.service;

import com.excel.dto.ExcelColumn;
import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class GenericExcelService<T> {
    public <T> ByteArrayInputStream generateExcel(List<T> data, List<ExcelColumn> columns, String sheetName) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet(sheetName);

            // Create header row with styling
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = createHeaderStyle(workbook);

            // Set up column headers
            for (int i = 0; i < columns.size(); i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columns.get(i).getHeader());
                cell.setCellStyle(headerStyle);

                // Set column width if specified
                if (columns.get(i).getWidth() > 0) {
                    sheet.setColumnWidth(i, columns.get(i).getWidth() * 256); // 256 = 1 character width
                }
            }

            // Fill data rows
            int rowIndex = 1;
            for (T item : data) {
                Row row = sheet.createRow(rowIndex++);

                for (int i = 0; i < columns.size(); i++) {
                    Cell cell = row.createCell(i);
                    String fieldName = columns.get(i).getFieldName();
                    Object value = getFieldValue(item, fieldName);
                    setCellValue(cell, value);
                }
            }

            // Auto-size columns that don't have a specified width
            for (int i = 0; i < columns.size(); i++) {
                if (columns.get(i).getWidth() <= 0) {
                    sheet.autoSizeColumn(i);
                }
            }

            // Write to output stream
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            return new ByteArrayInputStream(outputStream.toByteArray());
        }
    }
    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerStyle.setFont(headerFont);

        return headerStyle;
    }


    private Object getFieldValue(Object object, String fieldName) {
        try {
            // First try using getter method (getBeanProperty)
            String getterName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
            try {
                Method getter = object.getClass().getMethod(getterName);
                return getter.invoke(object);
            } catch (NoSuchMethodException e) {
                // If getter not found, try accessing field directly
                Field field = object.getClass().getDeclaredField(fieldName);
                field.setAccessible(true);
                return field.get(object);
            }
        } catch (Exception e) {
            return null; // Return null if field not found or accessible
        }
    }
    private void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setCellValue("");
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Integer || value instanceof Long) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Double || value instanceof Float) {
            cell.setCellValue(((Number) value).doubleValue());
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue(value.toString());
        }
    }
}