package com.istl.springapachepoi.controller;

import com.istl.springapachepoi.service.EmployeeXLGeneratorService;
import lombok.NoArgsConstructor;
import org.springframework.beans.factory.annotation.*;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/v1/employee")
@NoArgsConstructor
public class EmployeeController {

    @Autowired
    private EmployeeXLGeneratorService employeeXlGeneratorService;

    @GetMapping("/genXL")
    public void genXL() throws Exception
    {
        employeeXlGeneratorService.write("Employees");
    }
}
