package com.dao;

import com.models.Employee;
import com.models.EmployeeDetails;
import com.util.Constants;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Executable;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import static com.util.Constants.FILE_NAME;
import static com.util.Constants.LID_DETAILS_XLSX;
import static com.util.Constants.columnHeaders;

@Component
public class ExcelReadWrite {
    /* // Initializing employees data to insert into the excel file
    static {
        Calendar dateOfBirth = Calendar.getInstance();
        dateOfBirth.set(1992, 7, 21);
        employees.add(new Employee("Rajeev Singh", "rajeev@example.com",
                dateOfBirth.getTime(), 1200000.0));

        dateOfBirth.set(1965, 10, 15);
        employees.add(new Employee("Thomas cook", "thomas@example.com",
                dateOfBirth.getTime(), 1500000.0));

        dateOfBirth.set(1987, 4, 18);
        employees.add(new Employee("Steve Maiden", "steve@example.com",
                dateOfBirth.getTime(), 1800000.0));
    }*/

    public void write(Employee employee) throws IOException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* *//* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way *//*
        CreationHelper createHelper = workbook.getCreationHelper();*/

        // Create a Sheet
        Sheet sheet = workbook.createSheet(employee.getDay());

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create header cells
        for (int i = 0; i < columnHeaders.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columnHeaders[i]);
            cell.setCellStyle(headerCellStyle);
        }

       /* // Create Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));*/

        // Create first entry after header in the sheet
        Row row = sheet.createRow(1);
        row.createCell(0).setCellValue(employee.getLid());
        Cell loginTimeCell = row.createCell(1);
        loginTimeCell.setCellValue(employee.getLoginTime());
        loginTimeCell.setCellType(CellType.STRING);

        Cell logoffTimeCell = row.createCell(2);
        logoffTimeCell.setCellType(CellType.STRING);
        logoffTimeCell.setCellValue(employee.getLogoffTime());

        // Resize all columns to fit the content size
        for (int i = 0; i < columnHeaders.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = null;

        fileOut = new FileOutputStream(FILE_NAME);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

    public void update(Employee employee, String eventType) throws IOException, InvalidFormatException {
        // Obtain a workbook from the excel file
        FileInputStream file = new FileInputStream(FILE_NAME);
        Workbook workbook = WorkbookFactory.create(file);

        if (Objects.nonNull(workbook)) {
            // Get Sheet at index 0
            boolean lidExists = false;
            Sheet sheet = workbook.getSheet(employee.getDay());
            if (Objects.nonNull(sheet)) {
                for (Row myrow : sheet) {
                    Cell lidCell = myrow.getCell(0);
                    lidCell.setCellType(CellType.STRING);
                    if (employee.getLid().equalsIgnoreCase(String.valueOf(lidCell))) {
                        lidExists = true;
                        if ("login".equalsIgnoreCase(eventType)) {
                            Cell cell = myrow.getCell(1);
                            if (cell == null) {
                                cell = myrow.createCell(1);
                                cell.setCellType(CellType.STRING);
                                cell.setCellValue(employee.getLoginTime());
                            }
                            break;
                        } else if ("logoff".equalsIgnoreCase(eventType)) {
                            Cell cell = myrow.getCell(2);
                            if (cell == null)
                                cell = myrow.createCell(2);
                            cell.setCellType(CellType.STRING);
                            cell.setCellValue(employee.getLogoffTime());
                            break;
                        }
                    }
                }
            } else {
                // Create a Sheet
                sheet = workbook.createSheet(employee.getDay());

                // Create a Font for styling header cells
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                headerFont.setFontHeightInPoints((short) 14);
                headerFont.setColor(IndexedColors.RED.getIndex());

                // Create a CellStyle with the font
                CellStyle headerCellStyle = workbook.createCellStyle();
                headerCellStyle.setFont(headerFont);

                // Create a Row
                Row headerRow = sheet.createRow(0);

                // Create header cells
                for (int i = 0; i < columnHeaders.length; i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(columnHeaders[i]);
                    cell.setCellStyle(headerCellStyle);
                    
                }
            }

            if (!lidExists) {
                int rowNum = sheet.getPhysicalNumberOfRows();
                Row row = sheet.createRow(rowNum);

                row.createCell(0).setCellValue(employee.getLid());

                Cell loginTimeCell = row.createCell(1);
                loginTimeCell.setCellValue(employee.getLoginTime());
                loginTimeCell.setCellType(CellType.STRING);

                Cell logoffTimeCell = row.createCell(2);
                logoffTimeCell.setCellType(CellType.STRING);
                logoffTimeCell.setCellValue(employee.getLogoffTime());

                // Resize all columns to fit the content size
                for (int i = 0; i < columnHeaders.length; i++) {
                    sheet.autoSizeColumn(i);
                }

            }
            // Write the output to the file
            FileOutputStream fileOut = null;
            file.close();
            fileOut = new FileOutputStream(FILE_NAME);
            workbook.write(fileOut);
            fileOut.close();

            // Closing the workbook
            workbook.close();
        }
    }

    public List<EmployeeDetails> readLidDetails() throws IOException, InvalidFormatException {
        List<EmployeeDetails> employeeDetailsList = new ArrayList<>();
        // Obtain a workbook from the excel file
        FileInputStream file = new FileInputStream(LID_DETAILS_XLSX);
        Workbook workbook = WorkbookFactory.create(file);
        if (Objects.nonNull(workbook)) {
            // Get Sheet at index 0
            Sheet sheet = workbook.getSheetAt(0);
            if (Objects.nonNull(sheet)) {
                int index = 0;
                for (Row myrow : sheet) {
                    if (index!=0) {
                        EmployeeDetails employeeDetails = new EmployeeDetails();
                        Cell lidCell = myrow.getCell(0);
                        lidCell.setCellType(CellType.STRING);
                        employeeDetails.setLid(lidCell.getStringCellValue());
                        Cell nameCell = myrow.getCell(1);
                        nameCell.setCellType(CellType.STRING);
                        employeeDetails.setName(nameCell.getStringCellValue());
                        Cell emailCell = myrow.getCell(2);
                        emailCell.setCellType(CellType.STRING);
                        employeeDetails.setEmail(emailCell.getStringCellValue());
                        employeeDetailsList.add(employeeDetails);
                    }else {
                        index++;
                    }
                }
            }
        }
        return employeeDetailsList;
    }

	public void writeSprintDetails(String lid, String sprintEmailid, String featureTeamName, String projectName,
			String sprintNumber, String a1, String a2,String a3, String b1, String b2, String b3, String c1,String c2, String c3, String d1,String d2, String d3) throws IOException, EncryptedDocumentException, InvalidFormatException {
		 // Create a Workbook
		FileInputStream file = new FileInputStream(Constants.FILE_NAME);
        Workbook workbook = WorkbookFactory.create(file);
        if (Objects.nonNull(workbook)) {
            // Get Sheet at index 0
            boolean lidExists = false;
            Sheet sheet = workbook.getSheetAt(0);
            if (Objects.nonNull(sheet)) {
            	int rownum = sheet.getLastRowNum();
            	Row myrow = sheet.createRow(++rownum);
                   Cell lidCell = myrow.getCell(0);
                	   lidCell=myrow.createCell(0);
                	   lidCell.setCellType(CellType.STRING);
                	   lidCell.setCellValue(lid);
                	   myrow.createCell(1).setCellValue(sprintEmailid);
                	   myrow.createCell(2).setCellValue(featureTeamName);
                	   myrow.createCell(3).setCellValue(projectName);
                	   myrow.createCell(4).setCellValue(sprintNumber);
                	   myrow.createCell(5).setCellValue(a1);
                	   myrow.createCell(6).setCellValue(b1);
                	   myrow.createCell(7).setCellValue(c1);
                	   myrow.createCell(8).setCellValue(d1);
                	   myrow = sheet.createRow(++rownum);
                	   myrow.createCell(5).setCellValue(a2);
                	   myrow.createCell(6).setCellValue(b2);
                	   myrow.createCell(7).setCellValue(c2);
                	   myrow.createCell(8).setCellValue(d2);
                	   myrow = sheet.createRow(++rownum);
                	   myrow.createCell(5).setCellValue(a3);
                	   myrow.createCell(6).setCellValue(b3);
                	   myrow.createCell(7).setCellValue(c3);
                	   myrow.createCell(8).setCellValue(d3);
                 
                  
                
            } 
        }
            // Write the output to the file
            FileOutputStream fileOut = null;
            file.close();
            fileOut = new FileOutputStream(FILE_NAME);
            workbook.write(fileOut);
            fileOut.close();

            // Closing the workbook
            workbook.close();
        }
        
		
}