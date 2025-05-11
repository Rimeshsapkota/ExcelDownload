package com.excel.service;

import org.apache.poi.ss.formula.functions.T;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Service
public class EmployeeServiceImpl {
    @Autowired
    private GenericExcelService genericExcelService;

    public List<Dto.Employee<T>> employeeList() throws IOException {
        Dto.Employee<T> employee = new Dto.Employee();
        List<Dto.Employee<T>> employeeList = new ArrayList<>();
        employee.setFullName("pratik regmi");
        employee.setSalary(1000);
        employeeList.add(employee);
        genericExcelService.generateExcel(employeeList, Dto.Employee.class);
        return employeeList;
    }
}
