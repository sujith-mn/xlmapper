package com.xlm;

import com.xlm.util.XLColumn;
import com.xlm.util.XLSheet;

@XLSheet(name="Sheet1")
public class Employee {
    @XLColumn(name="Name")
    private String name;
    @XLColumn(name="EmpId")
    private Integer empId;
    @XLColumn(name="Grade")
    private String grade;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getEmpId() {
        return empId;
    }

    public void setEmpId(Integer empId) {
        this.empId = empId;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "name='" + name + '\'' +
                ", empId=" + empId +
                ", grade='" + grade + '\'' +
                '}';
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }


}
