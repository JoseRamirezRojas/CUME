package com.ciencias.CUME.api.service;

import java.io.IOException;
import java.lang.reflect.Field;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ciencias.CUME.api.model.PhysChemData;
 
public class PhysChemService {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet;
    private PhysChemData physChemData;
     
    public PhysChemService(PhysChemData physChemData) {
        this.physChemData = physChemData;
        workbook = new XSSFWorkbook();
        sheet = workbook.createSheet("Calidad fisicoquímica");
    }
 
 
    private void writeHeaderLine() {   
        // merging cells for sheet title 
        sheet.addMergedRegion(new CellRangeAddress(0,0, 0, 5));
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        font.setBold(true);
        font.setFontHeight(14);
        style.setFont(font);  // style for header title

        Row row = sheet.createRow(0);
        createCell(row,0,
         "Registro de parámetros fisicoquímicos de ríos de la Cuenca de México.", style);   
    }

    private void writeAuthorLines() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFillPattern(FillPatternType.NO_FILL); // reset to no background
        font.setBold(true);
        font.setFontHeight(12);
        style.setFont(font);  //reset font
        
        Row row = sheet.createRow(1);
        createCell(row ,0, "NOMBRE DEL PROYECTO:", style); 
        row = sheet.createRow(2);     
        createCell(row,0, "Nombre de la cuenca y subcuenca:", style); 
        createCell(row,2, "Fecha:", style); 
        row = sheet.createRow(3);   
        createCell(row,0, "Localidad:", style);
        createCell(row,2, "Hora:", style);
        row = sheet.createRow(4);   
        createCell(row,0, "Altitud:", style);
        row = sheet.createRow(5);   
        createCell(row,0, "Completaron la forma: (nombres)", style);
    }

    private void writeDataHeaderLine() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);          

        Row row = sheet.createRow(7);        
        sheet.addMergedRegion(new CellRangeAddress(7,8, 0, 1));
        createCenteredCell(row, 0, "Parámetros del agua", style);

        sheet.addMergedRegion(new CellRangeAddress(7,8, 2, 2));
        sheet.addMergedRegion(new CellRangeAddress(7,8, 3, 3));
        sheet.addMergedRegion(new CellRangeAddress(7,8, 4, 4));
        sheet.addMergedRegion(new CellRangeAddress(7,8, 5, 5));
        createCenteredCell(row, 2, "Prueba 1", style);
        createCenteredCell(row, 3, "Prueba 2", style);
        createCenteredCell(row, 4, "Prueba 3", style);
        createCenteredCell(row, 5, "Promedio", style);

        int numMerged = sheet.getNumMergedRegions();
        for(int i= 1; i<numMerged;i++){  // apply border to merged cells
            CellRangeAddress mergedRegions = sheet.getMergedRegion(i);
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, mergedRegions, sheet);
            RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), mergedRegions, sheet);
        }
    }

    private void writeDataLines() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);

        sheet.addMergedRegion(new CellRangeAddress(9,12, 0, 0));
        Row row = sheet.createRow(9);
        createCenteredCell(row, 0, "Físicos", style);
        
        createCell(row, 1, "Temperatura del agua (°C)", style);
        createCell(row, 2, physChemData.getTemperatura1(), style);
        createCell(row, 3, physChemData.getTemperatura2(), style);
        createCell(row, 4, physChemData.getTemperatura3(), style);
        createFormulaCell( row, 5, "AVERAGE(C10:E10)",style); 
        
        
             
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

    private void createCenteredCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet.autoSizeColumn(columnCount);
        Cell cell = row.createCell(columnCount);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
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
     
    
     
    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeAuthorLines();
        writeDataHeaderLine();
        for(int i=0; i<6 ;i++){
            sheet.autoSizeColumn(i);  //adjust width of columns
        }
        writeDataLines();
        sheet.autoSizeColumn(1);
        
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.flush();
        outputStream.close();
         
    }

}