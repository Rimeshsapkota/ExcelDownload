package com.excel.service;

import com.excel.dto.ExcelColumn;
import com.excel.utility.ExcelExportable;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class StudentExcelExporterService implements ExcelExportable {
    @Override
    public Map<String, String> getExcelColumnMapping() {
        Map<String, String> columns = new LinkedHashMap<>();
        columns.put("ID", "id");
        columns.put("Student Name", "name");
        columns.put("Grade", "grade");
        columns.put("GPA", "gpa");
        columns.put("Department", "department");
        return columns;
    }

    public List<ExcelColumn> getExcelColumns() {
        List<ExcelColumn> columns = new ArrayList<>();
        columns.add(new ExcelColumn("ID", "id", 5));
        columns.add(new ExcelColumn("Student Name", "name", 20));
        columns.add(new ExcelColumn("Grade", "grade", 10));
        columns.add(new ExcelColumn("GPA", "gpa", 8));
        columns.add(new ExcelColumn("Department", "department", 15));
        return columns;
    }

    public String getSheetName() {
        return "Students";
    }


}
