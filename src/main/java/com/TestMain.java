package com;

import com.dao.ExcelReadWrite;
import com.models.Employee;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.util.Arrays;
import java.util.Base64;

public class TestMain {
    public static void main1(String[] args) throws IOException, InvalidFormatException {
        ExcelReadWrite excelReadWrite = new ExcelReadWrite();

        Employee employee1 = new Employee();
        employee1.setLid("123411");
        employee1.setLoginTime("11:15");
        employee1.setDay("11-04-2020");

        String eventType1 = "login";
        excelReadWrite.update(employee1, eventType1);

        Employee employee = new Employee();
        employee.setLid("123411");
        employee.setLogoffTime("12:30");
        employee.setDay("11-04-2020");

        String eventType = "logoff";
        excelReadWrite.update(employee, eventType);


        Employee employee2 = new Employee();
        employee2.setLid("123411");
        employee2.setLoginTime("01:20");
        employee2.setDay("11-04-2020");

        excelReadWrite.update(employee2, eventType1);

        Employee employee0 = new Employee();
        employee0.setLid("123411");
        employee0.setLogoffTime("1:30");
        employee0.setDay("11-04-2020");

        excelReadWrite.update(employee0, eventType);

        /*TestController testController = new TestController();
        testController.fileExists();*/
    }
}
