package com.excel.ExcelController;

import com.excel.dto.Student;
import com.excel.service.GenericExcelService;
import com.excel.service.StudentExcelExporterService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

@RestController
@RequestMapping("/api/excel")
public class ExcelDownloadController {
    @Autowired
    private GenericExcelService excelService;

    @Autowired
    private StudentExcelExporterService studentExcelExporter;


    @GetMapping("/students")
    public ResponseEntity<InputStreamResource> downloadStudentExcel() throws IOException {
        // Sample student data - in a real app, this would come from a database
        List<Student> students = Arrays.asList(
                new Student(1L, "John Smith", "A", 3.8, "Computer Science"),
                new Student(2L, "Emma Johnson", "B+", 3.6, "Mathematics"),
                new Student(3L, "Michael Brown", "A-", 3.7, "Physics"),
                new Student(4L, "Sophia Williams", "B", 3.2, "Chemistry"),
                new Student(5L, "William Davis", "A+", 4.0, "Engineering")
        );

        ByteArrayInputStream excelFile = excelService.generateExcel(
                students,
                studentExcelExporter.getExcelColumns(),
                studentExcelExporter.getSheetName()
        );

        return getExcelResponse(excelFile, "students.xlsx");
    }
    private ResponseEntity<InputStreamResource> getExcelResponse(ByteArrayInputStream excelFile, String filename) {
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=" + filename);

        return ResponseEntity
                .ok()
                .headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(new InputStreamResource(excelFile));
    }
}
