package com.models;


import javax.validation.constraints.NotBlank;
import javax.validation.constraints.Pattern;
import javax.validation.constraints.Size;

public class Employee {
    @NotBlank
    private String lid;
    private String loginTime;
    private String logoffTime;
    private String day;

    public Employee() {
    }

    public Employee(String lid, String loginTime, String logoffTime) {
        this.lid = lid;
        this.loginTime = loginTime;
        this.logoffTime = logoffTime;
    }
    public String getLid() {
        return lid;
    }

    public void setLid(String lid) {
        this.lid = lid;
    }

    public String getLoginTime() {
        return loginTime;
    }

    public void setLoginTime(String loginTime) {
        this.loginTime = loginTime;
    }

    public String getLogoffTime() {
        return logoffTime;
    }

    public void setLogoffTime(String logoffTime) {
        this.logoffTime = logoffTime;
    }

    public String getDay() {
        return day;
    }

    public void setDay(String day) {
        this.day = day;
    }

    @Override
    public String toString() {
        return "Employee{" +
                "lid='" + lid + '\'' +
                ", loginTime='" + loginTime + '\'' +
                ", logoffTime='" + logoffTime + '\'' +
                ", day='" + day + '\'' +
                '}';
    }
}
