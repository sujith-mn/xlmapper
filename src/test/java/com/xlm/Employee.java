package com.xlm;

import com.xlm.util.XLColumn;
import com.xlm.util.XLSheet;

import java.util.Date;

@XLSheet(name="Sheet1")
public class Employee {
    @XLColumn(name="Name")
    private String name;
    @XLColumn(name="EmpId")
    private double empId;
    @XLColumn(name="Grade")
    private String grade;

    @XLColumn(name="dob")
    private Date dob;

    @XLColumn(name="active")
    private boolean active;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public double getEmpId() {
        return empId;
    }

    public void setEmpId(double empId) {
        this.empId = empId;
    }

    public boolean isActive() {
        return active;
    }

    public void setActive(boolean active) {
        this.active = active;
    }

    public Date getDob() {
        return dob;
    }

    public void setDob(Date dob) {
        this.dob = dob;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "name='" + name + '\'' +
                ", empId=" + empId +
                ", grade='" + grade + '\'' +
                ",dob='"+dob  + '\'' +
                ",active='"+active  + '\'' +
                '}';
    }

    public String getGrade() {
        return grade;
    }

    public void setGrade(String grade) {
        this.grade = grade;
    }


}