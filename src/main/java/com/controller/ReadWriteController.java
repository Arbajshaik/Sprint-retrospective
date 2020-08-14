package com.controller;

import com.dao.ExcelReadWrite;
import com.models.EmployeeDetails;
import com.util.Constants;
import com.models.Employee;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.NotNull;
import javax.validation.constraints.Pattern;
import javax.validation.constraints.Size;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.List;

@Controller
public class ReadWriteController {
    Logger logger = LoggerFactory.getLogger(ReadWriteController.class);
    @Autowired
    ExcelReadWrite excelReadWrite;
    int count = 0;

    @GetMapping("logTime")
    public String logTime(@RequestParam @NotNull(message = "LID is mandatory") @Pattern(regexp =
            "^L[0-9]*$") @Size(min= 7, max = 7, message = "LID Should be valid length") String lid, @RequestParam
            String time, @RequestParam String eventType) {
        DateFormat dayFormatter = new SimpleDateFormat("dd-MM-yyyy");
        DateFormat timeFormatter = new SimpleDateFormat("hh:mm");
        long milliSeconds = Long.parseLong(time);
        System.out.println(milliSeconds);
        Calendar calendar = Calendar.getInstance();
        calendar.setTimeInMillis(milliSeconds);
        System.out.println(String.valueOf(dayFormatter.format(calendar.getTime())));
        System.out.println(String.valueOf(timeFormatter.format(calendar.getTime())));
        Employee emp = new Employee();
        emp.setLid(lid);
        emp.setDay(String.valueOf(dayFormatter.format(calendar.getTime())));
        if (eventType.equalsIgnoreCase("login")) {
            emp.setLoginTime(String.valueOf(timeFormatter.format(calendar.getTime())));
        } else if (eventType.equalsIgnoreCase("logoff")) {
            emp.setLogoffTime(String.valueOf(timeFormatter.format(calendar.getTime())));
        }
        logger.info("Employee Object: {}", emp);
        try {
            if (fileExists()) {
                excelReadWrite.update(emp, eventType);
            } else {
                excelReadWrite.write(emp);
            }
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
            return "error";
        }
        return "success";
    }

    @GetMapping("getAllLids")
    @ResponseBody
    public List<EmployeeDetails> getAllLids(){
        try {
            return excelReadWrite.readLidDetails();
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
            logger.info("Unable to fetch LID data");
            return null;
        }
    }

    public boolean fileExists() {
        File file = new File(Constants.FILE_NAME);
        logger.info("Files exists? {}", file.exists());
        return file.exists();
    }
    
    @GetMapping("sprintDetails")
    public void sprintTracker(HttpServletResponse response,@RequestParam String lid, @RequestParam String sprintEmailid, @RequestParam String featureTeamName, @RequestParam String projectName, @RequestParam String sprintNumber,
    		@RequestParam String a1, @RequestParam String a2, @RequestParam String a3, @RequestParam String b1, @RequestParam String b2, @RequestParam String b3, @RequestParam String c1, @RequestParam String c2, @RequestParam String c3,
    		@RequestParam String d1, @RequestParam String d2, @RequestParam String d3) throws EncryptedDocumentException, InvalidFormatException, IOException {
    	File file = new File(Constants.FILE_NAME);
        System.out.println("lid"+ lid+" email"+sprintEmailid);
    	try {
            if (file.exists()) {
                excelReadWrite.writeSprintDetails(lid, sprintEmailid, featureTeamName, projectName, sprintNumber, a1,a2,a3,b1,b2,b3,c1,c2,c3,d1,d2,d3);
            }
        } catch (IOException e) {
            e.printStackTrace();

        }
    	 byte[] bytes =Files.readAllBytes(file.toPath());
         response.setContentType("application/xlsx");
         response.setHeader("Content-Disposition", "inline; filename=SpringDetails.xlsx");
         response.setContentLength(bytes.length);
         try {
             response.getOutputStream().write(bytes);
             response.getOutputStream().flush();
         } catch (IOException e) {
             logger.error("IOException", e);
         }
    }

}
