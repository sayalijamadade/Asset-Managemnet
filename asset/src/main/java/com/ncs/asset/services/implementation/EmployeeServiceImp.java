package com.ncs.asset.services.implementation;

import com.ncs.asset.model.Employee;
import com.ncs.asset.repository.EmployeeRepository;
import com.ncs.asset.services.interfaces.EmployeeService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

@Service
public class EmployeeServiceImp implements EmployeeService {

    @Autowired
    EmployeeRepository employeeRepository;

    @Override
    public Employee saveEmployee(Employee employee) {
        return employeeRepository.save(employee);
    }
}
