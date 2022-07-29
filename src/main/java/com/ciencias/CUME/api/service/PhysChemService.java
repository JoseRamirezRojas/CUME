package com.ciencias.CUME.api.service;

import java.io.IOException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
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
        createCenteredCell(row, 0, "Físicos", style);  //physical parameters
        createCell(row, 1, "Temperatura del agua (°C)", style);
        createCell(row, 2, physChemData.getTemperatura1(), style);
        createCell(row, 3, physChemData.getTemperatura2(), style);
        createCell(row, 4, physChemData.getTemperatura3(), style);
        createFormulaCell( row, 5, "AVERAGE(C10:E10)",style); 
        row = sheet.createRow(10);     
        createCell(row, 1, "Conductividad Específica (μS)", style);
        createCell(row, 2, physChemData.getConductividad1(), style);
        createCell(row, 3, physChemData.getConductividad2(), style);
        createCell(row, 4, physChemData.getConductividad3(), style);
        createFormulaCell( row, 5, "AVERAGE(C11:E11)",style); 
        row = sheet.createRow(11); 
        createCell(row, 1, "Oxígeno disuelto (%)", style);
        createCell(row, 2, physChemData.getOxiDisuelto1(), style);
        createCell(row, 3, physChemData.getOxiDisuelto2(), style);
        createCell(row, 4, physChemData.getOxiDisuelto3(), style);
        createFormulaCell( row, 5, "AVERAGE(C12:E12)",style); 
        row = sheet.createRow(12); 
        createCell(row, 1, "Oxígeno solubre (mg/L)", style);
        createCell(row, 2, physChemData.getOxiSolubre1(), style);
        createCell(row, 3, physChemData.getOxiSolubre2(), style);
        createCell(row, 4, physChemData.getOxiSolubre3(), style);
        createFormulaCell( row, 5, "AVERAGE(C13:E13)",style); 
        
        sheet.addMergedRegion(new CellRangeAddress(13,17, 0, 0));
        row = sheet.createRow(13);
        createCenteredCell(row, 0, "Químicos", style);   //chemical parameters
        createCell(row, 1, "pH", style);
        createCell(row, 2, physChemData.getPH1(), style);
        createCell(row, 3, physChemData.getPH2(), style);
        createCell(row, 4, physChemData.getPH3(), style);
        createFormulaCell( row, 5, "AVERAGE(C14:E14)",style); 
        row = sheet.createRow(14); 
        createCell(row, 1, "Fósforo (mg/L)", style);
        createCell(row, 2, physChemData.getFosforo1(), style);
        createCell(row, 3, physChemData.getFosforo2(), style);
        createCell(row, 4, physChemData.getFosforo3(), style);
        createFormulaCell( row, 5, "AVERAGE(C15:E15)",style);
        row = sheet.createRow(15); 
        createCell(row, 1, "Nitrito (mg/L)", style);
        createCell(row, 2, physChemData.getNitrito1(), style);
        createCell(row, 3, physChemData.getNitrito2(), style);
        createCell(row, 4, physChemData.getNitrito3(), style);
        createFormulaCell( row, 5, "AVERAGE(C16:E16)",style); 
        row = sheet.createRow(16); 
        createCell(row, 1, "Nitrato (mg/L)", style);
        createCell(row, 2, physChemData.getNitrato1(), style);
        createCell(row, 3, physChemData.getNitrato2(), style);
        createCell(row, 4, physChemData.getNitrato3(), style);
        createFormulaCell( row, 5, "AVERAGE(C17:E17)",style); 
        row = sheet.createRow(17); 
        createCell(row, 1, "Amonio (mg/L)", style);
        createCell(row, 2, physChemData.getAmonio1(), style);
        createCell(row, 3, physChemData.getAmonio2(), style);
        createCell(row, 4, physChemData.getAmonio3(), style);
        createFormulaCell( row, 5, "AVERAGE(C18:E18)",style); 

        sheet.addMergedRegion(new CellRangeAddress(18,22, 0, 0));
        row = sheet.createRow(18);
        createCenteredCell(row, 0, "Sustrato inorgánico (% en el área muestreada)", 
          style);   //sustratos

        createCell(row, 1, "Rocas (>256 mm)", style);
        createCell(row, 2, physChemData.getRocas1(), style);
        createCell(row, 3, physChemData.getRocas2(), style);
        createCell(row, 4, physChemData.getRocas3(), style);
        createFormulaCell( row, 5, "AVERAGE(C19:E19)",style); 
        row = sheet.createRow(19); 
        createCell(row, 1, "Canto (64-256 mm)", style);
        createCell(row, 2, physChemData.getCanto1(), style);
        createCell(row, 3, physChemData.getCanto2(), style);
        createCell(row, 4, physChemData.getCanto3(), style);
        createFormulaCell( row, 5, "AVERAGE(C20:E20)",style); 
        row = sheet.createRow(20); 
        createCell(row, 1, "Grava (2-64 mm)", style);
        createCell(row, 2, physChemData.getGrava1(), style);
        createCell(row, 3, physChemData.getGrava2(), style);
        createCell(row, 4, physChemData.getGrava3(), style);
        createFormulaCell( row, 5, "AVERAGE(C21:E21)",style); 
        row = sheet.createRow(21); 
        createCell(row, 1, "Arena (0.06-2 mm)", style);
        createCell(row, 2, physChemData.getArena1(), style);
        createCell(row, 3, physChemData.getArena2(), style);
        createCell(row, 4, physChemData.getArena3(), style);
        createFormulaCell( row, 5, "AVERAGE(C22:E22)",style); 
        row = sheet.createRow(22); 
        createCell(row, 1, "Arcilla (0.004 mm)", style);
        createCell(row, 2, physChemData.getArcilla1(), style);
        createCell(row, 3, physChemData.getArcilla2(), style);
        createCell(row, 4, physChemData.getArcilla3(), style);
        createFormulaCell( row, 5, "AVERAGE(C23:E23)",style);
        
        //aforo
        sheet.addMergedRegion(new CellRangeAddress(24,34, 0, 0));
        row = sheet.createRow(24);
        createCenteredCell(row, 0, "Aforo (㎥/s)", style);  
        createCell(row, 1, "Ancho del río (m):", style);
        row = sheet.createRow(25);
        createCell(row, 1, "Distancia entre intervalos (m):", style);
        sheet.autoSizeColumn(1);
        for(int i = 1; i <= 9; i++){
            row = sheet.createRow(25 + i);
            createCell(row, 1, "Prof. " + i + ":", style);
            createCell(row, 3, "Vel. Corriente " + i+":", style);
            createCell(row, 5, "Q" + i+":", style);
            if(i == 1 || i == 9)
                createFormulaCell(row,6,"(C"+(26+i)+"*C26/2)*E"+(26+i),style); // B*h/2 * V
            else
                createFormulaCell(row,6,"((C"+(25+i)+"+C"+(26+i)+")*C26/2)*E"+(26+i),  // ((B+b)*h/2) * V
                    style); 
        }
        sheet.autoSizeColumn(3);
        row = sheet.createRow(35);
        createCell(row, 5, "ΣQ (㎥/s) =", style);
        sheet.autoSizeColumn(5);
        createFormulaCell(row,6,"SUM(G27:G35)",style); // B*h/2 * V
    }
     
    private void createCell(Row row, int columnCount, Object value, CellStyle style) {
        Cell cell = row.createCell(columnCount);
        DataFormat format = workbook.createDataFormat();
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Float) {
            cell.setCellValue((Float) value);
            style.setDataFormat(format.getFormat("0.00"));
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
        DataFormat format = workbook.createDataFormat();
        Cell cell = row.createCell(columnCount);
        style.setDataFormat(format.getFormat("0.00"));
        cell.setCellFormula(formula);
        cell.setCellStyle(style);
    }
       
    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeAuthorLines();
        writeDataHeaderLine();
        writeDataLines();
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.flush();
        outputStream.close();    
    }

}