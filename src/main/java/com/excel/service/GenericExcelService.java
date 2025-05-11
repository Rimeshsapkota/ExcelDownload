package com.excel.service;

import org.apache.poi.ss.usermodel.*;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class GenericExcelService<T>{
    public ByteArrayInputStream generateExcel(List<T> data,Class<T> dat) throws IOException {
        try (Workbook workbook = new XSSFWorkbook ()) {
            Sheet sheet = workbook.createSheet("Employees");

            // Create header row
            Row headerRow = sheet.createRow(0);

            // Apply header style
            CellStyle headerStyle = workbook.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            Font headerFont = workbook.createFont();
            headerFont.setBold(true);
            headerFont.setColor(IndexedColors.WHITE.getIndex());
            headerStyle.setFont(headerFont);

            // Create header cells
            Cell headerCell = headerRow.createCell(0);
            headerCell.setCellValue("ID");
            headerCell.setCellStyle(headerStyle);

            headerCell = headerRow.createCell(1);
            headerCell.setCellValue("Name");
            headerCell.setCellStyle(headerStyle);

            headerCell = headerRow.createCell(2);
            headerCell.setCellValue("Department");
            headerCell.setCellStyle(headerStyle);

            headerCell = headerRow.createCell(3);
            headerCell.setCellValue("Salary");
            headerCell.setCellStyle(headerStyle);

            // Create data rows
            int rowIndex = 1;
            for (T employee  : data) {
                Row row = sheet.createRow(rowIndex++);

                Cell cell = row.createCell(0);
                cell.setCellValue(employee.getId());

                cell = row.createCell(1);
                cell.setCellValue(employee.getName());

                cell = row.createCell(2);
                cell.setCellValue(employee.getDepartment());

                cell = row.createCell(3);
                cell.setCellValue(employee.getSalary());
            }

            // Resize all columns to fit the content
            for (int i = 0; i < 4; i++) {
                sheet.autoSizeColumn(i);
            }

            // Write to ByteArrayOutputStream
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);

            return new ByteArrayInputStream(outputStream.toByteArray());
        }
    }

}
