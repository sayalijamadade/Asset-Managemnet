package com.ncs.asset.controller;

import com.ncs.asset.model.Employee;
import com.ncs.asset.model.Location;
import com.ncs.asset.services.interfaces.EmployeeService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping(value = "/employee")
public class EmployeeController {

    @Autowired
    EmployeeService employeeService;

    @PostMapping(value = "/save")
    public Employee persistEmployee(@RequestBody final Employee employee) {

        return employeeService.saveEmployee(employee);
    }

}
