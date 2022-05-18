package com.istl.springapachepoi.service;

import com.istl.springapachepoi.dto.Employee;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

@Service
public class EmployeeXLGeneratorService
{
    private static String[] columns = {"Name", "Email", "Date Of Birth", "Salary"};
    private static List<Employee> employees =  new ArrayList<>();

    private static String[][] datas = {
            {"Jobayedp8","jobayehhh@istlbd.com","15/01/1992","10000.00"},
            {"Jobayeda1","jobayeaaa@istlbd.com","15/01/1993","10001.00"},
            {"Jobayedb2","jobayeddd@istlbd.com","15/01/1994","10005.00"},
            {"Jobayedq3","jobayeppp@istlbd.com","15/01/1990","10002.00"}};

    public EmployeeXLGeneratorService(){
        this.setEmployeeData();
    }



    public Employee newEmployee(String name,String email,String date,double salary)
    {
        return new Employee(name, email,getDate(date), salary);
    }

    public Date getDate(String dateStr){
        String pattern = "dd/MM/yyyy";
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(pattern);
        Date date;
        try {
            date = simpleDateFormat.parse(dateStr);
        }
        catch (ParseException e){
            date = new Date();
        }
        return date;
    }

    private void setEmployeeData(){
        for (String[] data:datas)
        {
            employees.add(newEmployee(data[0],data[1],data[2],Double.parseDouble(data[3])) );
        }
    }

    public void write(String sheetName) throws Exception{
        Workbook workbook = new XSSFWorkbook();                         // Create Workbook
        CreationHelper createHelper = workbook.getCreationHelper();     //

        Sheet sheet = workbook.createSheet(sheetName);

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.BLACK.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setBorderTop(BorderStyle.THIN);
        headerCellStyle.setBorderLeft(BorderStyle.THIN);
        headerCellStyle.setBorderBottom(BorderStyle.DASHED);

        // Create a Row
        Row headerRow = sheet.createRow(0);
        sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, columns.length - 1));
        sheet.createFreezePane(0, 1);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);

            if(i==columns.length-1){
                headerCellStyle.setBorderRight(BorderStyle.THIN);
            }
            cell.setCellStyle(headerCellStyle);
        }

        // Create Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));

        // Create a CellStyle with border
        CellStyle borderCellStyle = workbook.createCellStyle();
        borderCellStyle.setBorderLeft(BorderStyle.THIN);
        borderCellStyle.setBorderBottom(BorderStyle.THIN);





        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(Employee employee: employees) {
            borderCellStyle.setBorderRight(BorderStyle.NONE);

            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(employee.getName());
            row.getCell(0).setCellStyle(borderCellStyle);

            row.createCell(1)
                    .setCellValue(employee.getEmail());
            row.getCell(1).setCellStyle(borderCellStyle);

            Cell dateOfBirthCell = row.createCell(2);
            dateOfBirthCell.setCellValue(employee.getDateOfBirth());
            dateOfBirthCell.setCellStyle(dateCellStyle);
            row.getCell(2).setCellStyle(borderCellStyle);

            borderCellStyle.setBorderRight(BorderStyle.THIN);
            row.createCell(3)
                    .setCellValue(employee.getSalary());
            row.getCell(3).setCellStyle(borderCellStyle);
        }


        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("poi-generated-file.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }



}
