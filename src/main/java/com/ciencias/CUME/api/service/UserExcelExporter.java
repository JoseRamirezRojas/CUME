package com.ciencias.CUME.api.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
 
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
 
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
 
public class UserExcelExporter {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
     
    public UserExcelExporter() {
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Users");
    }
 
 
    private void writeHeaderLine() {
        
         
        Row row = sheet.createRow(0);
         
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(16);
        style.setFont(font);
         
        createCell(row, 0, "User ID", style);      
        createCell(row, 1, "E-mail", style);       
        createCell(row, 2, "Full Name", style);    
        createCell(row, 3, "Roles", style);
        createCell(row, 4, "Enabled", style);
         
    }
     
    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else {
            cell.setCellValue((String) value);
        }
        cell.setCellStyle(style);
    }

    private void createFormulaCell(Row row, int columnCount, String formula, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        cell.setCellFormula(formula);
        cell.setCellStyle(style);
    }
     
    private void writeDataLines() {
        int rowCount = 1;
 
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(14);
        style.setFont(font);
                 
        // for (User user : listUsers) {
        Row row = sheet.createRow(rowCount++);
            // int columnCount = 0;
             
        createCell(row, 0, 1, style);
        createCell(row, 1, 2, style);
        createCell(row, 2, 3, style);
        createCell(row, 3, 4, style);
        createFormulaCell( row, 4, "AVERAGE(A2:D2)",style); 
             
        // }
    }
     
    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeDataLines();

        //response.reset();
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.flush();
        outputStream.close();
         
    }

    public ByteArrayInputStream prepareData()  {
        try(XSSFWorkbook workbook = new XSSFWorkbook() ){
            sheet = workbook.createSheet("data");

            Row row = sheet.createRow(0);
            CellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            // Creating header
            Cell cell = row.createCell(0);
            cell.setCellValue("First Name");
            cell.setCellStyle(headerCellStyle);

            cell = row.createCell(1);
            cell.setCellValue("Last Name");
            cell.setCellStyle(headerCellStyle);

            cell = row.createCell(2);
            cell.setCellValue("Mobile");
            cell.setCellStyle(headerCellStyle);

            cell = row.createCell(3);
            cell.setCellValue("Email");
            cell.setCellStyle(headerCellStyle);

            // Creating data rows for each customer
            for(int i = 0; i < 8; i++) {
                Row dataRow = sheet.createRow(i + 1);
                dataRow.createCell(0).setCellValue("FirstName " + i);
                dataRow.createCell(1).setCellValue("LastName " + i);
                dataRow.createCell(2).setCellValue("Mobile " + i);
                dataRow.createCell(3).setCellValue("Email " + i);
            }

            // Making size of column auto resize to fit with data
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);

            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            workbook.write(outputStream);
            workbook.close();
            return new ByteArrayInputStream(outputStream.toByteArray());
        } catch (IOException ex) {
            ex.printStackTrace();
            return null;
        }
    } 
}