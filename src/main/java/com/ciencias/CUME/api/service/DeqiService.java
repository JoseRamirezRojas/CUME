package com.ciencias.CUME.api.service;

import java.io.IOException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ciencias.CUME.api.model.DeqiData;

public class DeqiService {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet1;
    private XSSFSheet sheet2;
    private DeqiData deqiData;

    public DeqiService(DeqiData deqiData) {
        this.deqiData = deqiData;
        workbook = new XSSFWorkbook();
        sheet1 = workbook.createSheet("DEQI"); // <31 chars
        sheet2 = workbook.createSheet("Ficha de colecta diatomeas" ); // <31 chars
    }

    private void writeHeaderLine() {   
        // merging cells for sheet title 
        sheet1.addMergedRegion(new CellRangeAddress(0,0, 0, 5));
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        font.setBold(true);
        font.setFontHeight(14);
        style.setFont(font);  // style for header title

        Row row = sheet1.createRow(0);
        createCell(row,0,
            "Cálculo de índice de la calidad ecológica de diatomeas (DEQI) " +
            "en ríos de la Cuenca de México.", style);   
    }

    private void writeAuthorLines() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFillPattern(FillPatternType.NO_FILL); // reset to no background
        font.setBold(true);
        font.setFontHeight(12);
        style.setFont(font);  //reset font
        
        Row row = sheet1.createRow(1);
        createCell(row ,0, "NOMBRE DEL PROYECTO:", style); 
        createCell(row ,2, "Revisor:", style); 
        row = sheet1.createRow(2);     
        createCell(row,0, "Río:", style); 
        createCell(row,2, "Fecha de colecta:", style); 
        row = sheet1.createRow(3);   
        createCell(row,0, "Localidad:", style);
        createCell(row,2, "Fecha de revisión:", style);
        row = sheet1.createRow(4);   
        createCell(row,0, "Completaron la forma: (nombres)", style);
        row = sheet1.createRow(5);   
        createCell(row,0, "Código de la muestra", style);
        row = sheet1.createRow(6);   
        createCell(row,0, "Dilución", style);
        row = sheet1.createRow(7);   
        createCell(row,0, "DEQI", style);
        createFormulaCell(row,1, "G11", style); // TODO
        createCell(row,2, "Calidad ecológica", style);
        createFormulaCell(row,3, "H11", style);  // TODO
    }

    private void writeDataHeaderLine() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);

        Row row = sheet1.createRow(9);   
        createCell(row,2, "Valvas totales", style);
        createItalicCell(row,3, "Σ h");
        createItalicCell(row,4, "Σ hv");
        createCell(row,6, "DEQI", style);
        createCell(row,7, "Calidad", style);

        row = sheet1.createRow(10); 
        createFormulaCell(row,2,"SUM(C14:C500)",style);
        createFormulaCell(row,3,"SUM(D14:D500)",style);
        createFormulaCell(row,4,"SUM(E14:E500)",style);
        createFormulaCell(row,6,"E11/D11",style);
        // conditional formatting of DEQI
        SheetConditionalFormatting sheetCF = sheet1.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1  = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "1.5");
        // Condition 1: Cell Value Is LESSS than 1.5   (Blue Fill)
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 2: Cell Value Is  greater than 1.5  (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "1.5");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.LIME.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 3: Cell Value Is   greater than 2.5   (yellow Fill)
        ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "2.5");
        PatternFormatting fill3 = rule3.createPatternFormatting();
        fill3.setFillBackgroundColor(IndexedColors.YELLOW.index);
        fill3.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 4: Cell Value Is  GREATER than 3.5  (red Fill)
        ConditionalFormattingRule rule4 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "3.5");
        PatternFormatting fill4 = rule4.createPatternFormatting();
        fill4.setFillBackgroundColor(IndexedColors.RED.index);
        fill4.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 4: Cell Value Is  GREATER than 4.5  (red Fill)
        ConditionalFormattingRule rule5 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GT, "4.5");
        PatternFormatting fill5 = rule5.createPatternFormatting();
        fill5.setFillBackgroundColor(IndexedColors.DARK_RED.index);
        fill5.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        CellRangeAddress[] regions = { CellRangeAddress.valueOf("G11") };
        sheetCF.addConditionalFormatting(regions, rule1);
        sheetCF.addConditionalFormatting(regions, rule2);
        sheetCF.addConditionalFormatting(regions, rule3);
        sheetCF.addConditionalFormatting(regions, rule4);
        sheetCF.addConditionalFormatting(regions, rule5);

        // deqi from hyqiData
        CellReference cellReference = new CellReference("G11");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
        Row rowAux = sheet1.getRow(cellReference.getRow());
        Cell cell = rowAux.getCell(cellReference.getCol()); 
        CellValue cellValue = evaluator.evaluate(cell);     
        switch (cellValue.getCellType()) {
            case BOOLEAN:
                System.out.println(cell.getBooleanCellValue());
                break;
            case NUMERIC:
                double deqi = cell.getNumericCellValue();
                if (deqi <= 1.5)
                    createCenteredCell(row, 7,"Alta", style); 
                else if (deqi <= 2.5 )
                    createCenteredCell(row, 7,"Buena", style); 
                else if (deqi <= 3.5 )
                    createCenteredCell(row, 7,"Moderada", style);
                else if (deqi <= 4.5 )
                    createCenteredCell(row, 7,"Pobre", style);
                else if (deqi <= 5)
                    createCenteredCell(row, 7,"Mala", style);
                else
                    createCenteredCell(row, 7,"ERROR", style);
                break;
            case STRING:
                System.out.println(cell.getStringCellValue());
                break;
        }

        row = sheet1.createRow(11);        
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 0, 0));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 1, 1));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 2, 2));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 3, 3));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 4, 4));
        createItalicCell(row, 0, "v");
        createBoldCell(row, 1, "Especie");
        createBoldCell(row, 2, "Conteo");
        createItalicCell(row, 3, "h");
        createItalicCell(row, 4, "hv");
        // sheet1.autoSizeColumn(4);

        int numMerged = sheet1.getNumMergedRegions();
        for(int i= 6; i<numMerged;i++){  // apply border to merged cells
            CellRangeAddress mergedRegions = sheet1.getMergedRegion(i);
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, mergedRegions, sheet1);
            RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), mergedRegions, sheet1);
        }
    }


    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
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

    private void createItalicCell(Row row, int columnCount, Object value ) {
        Cell cell = row.createCell(columnCount);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(12);
        font.setItalic(true);
        style.setFont(font);
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

    private void createBoldCell(Row row, int columnCount, Object value ) {
        Cell cell = row.createCell(columnCount);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontHeight(12);
        style.setFont(font);
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

    private void createCenteredCell(Row row, int columnCount, Object value, CellStyle style) {
        sheet1.autoSizeColumn(columnCount);
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
        Cell cell = row.createCell(columnCount);
        cell.setCellFormula(formula);
        cell.setCellStyle(style);
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeAuthorLines();
        writeDataHeaderLine();
        //writeDataLines();
        //writeProtocolRef();
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.flush();
        outputStream.close();    
    }
}
