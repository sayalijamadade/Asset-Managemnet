package com.ncs.asset.repository;

import com.ncs.asset.model.Employee;
import com.ncs.asset.model.Location;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface EmployeeRepository  extends JpaRepository<Employee, Integer> {
}
