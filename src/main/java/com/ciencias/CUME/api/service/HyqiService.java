package com.ciencias.CUME.api.service;

import java.io.IOException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.ciencias.CUME.api.model.HyqiData;
 
public class HyqiService {
    private XSSFWorkbook workbook;
    private XSSFSheet sheet1;
    private XSSFSheet sheet2;
    private HyqiData hyqiData;
     
    public HyqiService(HyqiData hyqiData) {
        this.hyqiData = hyqiData;
        workbook = new XSSFWorkbook();
        sheet1 = workbook.createSheet("Calidad hidromofológica"); // <31 chars
        sheet2 = workbook.createSheet("Referencia protocolo"); // <31 chars
    }
 
    private void writeHeaderLine() {   
        // merging cells for sheet title 
        sheet1.addMergedRegion(new CellRangeAddress(0,0, 0, 4));
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        font.setBold(true);
        font.setFontHeight(14);
        style.setFont(font);  // style for header title

        Row row = sheet1.createRow(0);
        createCell(row,0,
         "Registro de parámetros hidromorfológicos de ríos de la Cuenca de México.", style);   
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
        row = sheet1.createRow(2);     
        createCell(row,0, "Nombre de la cuenca y subcuenca:", style); 
        createCell(row,3, "Fecha:", style); 
        row = sheet1.createRow(3);   
        createCell(row,0, "Localidad:", style);
        createCell(row,3, "Hora:", style);
        row = sheet1.createRow(4);   
        createCell(row,0, "Altitud:", style);
        row = sheet1.createRow(5);   
        createCell(row,0, "Completaron la forma: (nombres)", style);
        row = sheet1.createRow(6); 
        sheet1.addMergedRegion(new CellRangeAddress(6,6, 0,4 ));  
        createCell(row,0, "Vegetación de ribera", style);
        row = sheet1.createRow(7); 
        sheet1.addMergedRegion(new CellRangeAddress(7,8, 0, 0)); 
        createCenteredCell(row,0, "Tipo de bosque", style);
        createItalicCell(row,1, "Abies");
        createItalicCell(row,2, "Pinus");
        createCell(row,3, "Mixto", style);
        row = sheet1.createRow(9); 
        sheet1.addMergedRegion(new CellRangeAddress(9,9, 0, 4));  
        createCell(row,0, "Forma de vida dominante", style);
        row = sheet1.createRow(10); 
        sheet1.addMergedRegion(new CellRangeAddress(10,11, 0, 0)); 
        createCenteredCell(row,0, "Ribera derecha", style);
        createCell(row,1, "Árbol", style);
        createCell(row,2, "Arbusto", style);
        createCell(row,3, "Pasto", style);
        createCell(row,4, "Hierba", style);
        row = sheet1.createRow(12); 
        sheet1.addMergedRegion(new CellRangeAddress(12,13, 0, 0)); 
        createCenteredCell(row,0, "Ribera izquierda", style);
        createCell(row,1, "Árbol", style);
        createCell(row,2, "Arbusto", style);
        createCell(row,3, "Pasto", style);
        createCell(row,4, "Hierba", style);
    }

    private void writeDataHeaderLine() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);          

        Row row = sheet1.createRow(15);        
        sheet1.addMergedRegion(new CellRangeAddress(15,16, 0, 1));
        sheet1.addMergedRegion(new CellRangeAddress(15,16, 2, 2));
        sheet1.addMergedRegion(new CellRangeAddress(15,16, 3, 3));
        sheet1.addMergedRegion(new CellRangeAddress(15,16, 4, 4));
        createCenteredCell(row, 0, "Parámetros", style);
        createCenteredCell(row, 2, "Clasificación", style);
        createCenteredCell(row, 3, "Puntaje", style);
        createCenteredCell(row, 4, "Sección", style);
        sheet1.autoSizeColumn(4);

        int numMerged = sheet1.getNumMergedRegions();
        for(int i= 6; i<numMerged;i++){  // apply border to merged cells
            CellRangeAddress mergedRegions = sheet1.getMergedRegion(i);
            RegionUtil.setBorderBottom(BorderStyle.MEDIUM, mergedRegions, sheet1);
            RegionUtil.setBottomBorderColor(IndexedColors.BLACK.getIndex(), mergedRegions, sheet1);
        }
    }

    private void writeDataLines() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);

        sheet1.addMergedRegion(new CellRangeAddress(17,22, 0, 0));
        Row row = sheet1.createRow(17);
        createCenteredCell(row, 0, "Cuenca", style);  //basin parameters
        createCell(row, 1, "Cobertura vegetal: ribera derecha", style);
        // Add this when excel 2019 version is more popular
        // String switchClasif="SWITCH(D18,0,\"Pobre\",2,\"Malo\",3,\"Medio\",5,\"Óptimo\",\"ERROR\")";
        // createFormulaCell(row, 2, switchClasif, style);  
        getClasificacionParamsRiberas(row, hyqiData.getCoberturaDer(), style);
        createCell(row, 3, hyqiData.getCoberturaDer(), style);
        sheet1.addMergedRegion(new CellRangeAddress(17,22, 4, 4));
        createFormulaCell(row,4,"SUM(D18:D23)",style); //section sum  
        row = sheet1.createRow(18); 
        createCell(row, 1, "Cobertura vegetal: ribera izquierda", style);
        getClasificacionParamsRiberas(row, hyqiData.getCoberturaIzq(), style);
        createCell(row, 3, hyqiData.getCoberturaIzq(), style); 
        row = sheet1.createRow(19);
        createCell(row, 1, "Estabilidad del banco", style);
        getClasificacionParams(row, hyqiData.getEstabilidad(), style);
        createCell(row, 3, hyqiData.getEstabilidad(), style); 
        row = sheet1.createRow(20);
        createCell(row, 1, "Características del sustrato", style);
        getClasificacionParams(row, hyqiData.getSustrato(), style);
        createCell(row, 3, hyqiData.getSustrato(), style); 
        row = sheet1.createRow(21);
        createCell(row, 1, "Ganadería/agricultura: ribera derecha", style);
        getClasificacionParamsRiberas(row, hyqiData.getAgriculturaDer(), style);
        createCell(row, 3, hyqiData.getAgriculturaDer(), style); 
        row = sheet1.createRow(22);
        createCell(row, 1, "Ganadería/agricultura: ribera izquierda", style);
        getClasificacionParamsRiberas(row, hyqiData.getAgriculturaIzq(), style);
        createCell(row, 3, hyqiData.getAgriculturaIzq(), style);
        //createCell(row, 4,"Puntaje sección", style);
           
        
        sheet1.addMergedRegion(new CellRangeAddress(23,26, 0, 0));
        row = sheet1.createRow(23);
        createCenteredCell(row, 0, "Hidrología",style); //hidrology parameters
        createCell(row, 1, "Presencia de presas", style);
        if(hyqiData.getPresas() == 10)
            createCell(row, 2, "Óptimo", style);
        else if(hyqiData.getPresas() == 0)
            createCell(row, 2, "Pobre", style);
        else
            createCell(row, 2, "ERROR", style);
        createCell(row, 3, hyqiData.getPresas(), style); 
        sheet1.addMergedRegion(new CellRangeAddress(23,26, 4, 4));
        createFormulaCell(row,4,"SUM(D24:D27)",style); //section sum    
        row = sheet1.createRow(24);
        createCell(row, 1, "Regímenes velocidad/profundidad", style);
        getClasificacionParams(row, hyqiData.getRegimenes(), style);
        createCell(row, 3, hyqiData.getRegimenes(), style);
        row = sheet1.createRow(25);
        createCell(row, 1, "Alteración en el canal", style);
        getClasificacionParams(row, hyqiData.getCanal(), style);
        createCell(row, 3, hyqiData.getCanal(), style); 
        row = sheet1.createRow(26);
        createCell(row, 1, "Estado del canal", style);
        getClasificacionParams(row, hyqiData.getEstado(), style);
        createCell(row, 3, hyqiData.getEstado(), style);   
        //createCell(row, 4,"Puntaje sección", style);   
        //createFormulaCell(row,4,"SUM(D24:D27)",style); //section sum
        
        sheet1.addMergedRegion(new CellRangeAddress(27,30, 0, 0));
        row = sheet1.createRow(27);
        createCenteredCell(row, 0, "Perturbaciones antropogénicas",style); //antropogenic parameters
        createCell(row, 1, "Efluentes directos al río", style);
        if(hyqiData.getEfluentes() == 10)
            createCell(row, 2, "Óptimo", style);
        else if(hyqiData.getEfluentes() == 1)
            createCell(row, 2, "Pobre", style);
        else
            createCell(row, 2, "ERROR", style);
        createCell(row, 3, hyqiData.getEfluentes(), style); 
        sheet1.addMergedRegion(new CellRangeAddress(27,30, 4, 4));
        createFormulaCell(row,4,"SUM(D28:D31)",style); //section sum    
        row = sheet1.createRow(28);
        createCell(row, 1, "Desarrollo urbano", style);
        getClasificacionParams(row, hyqiData.getUrbano(), style);
        createCell(row, 3, hyqiData.getUrbano(), style);
        row = sheet1.createRow(29);
        createCell(row, 1, "Desarrollo humano", style);
        getClasificacionParams(row, hyqiData.getHumano(), style);
        createCell(row, 3, hyqiData.getHumano(), style); 
        row = sheet1.createRow(30);
        createCell(row, 1, "Presencia de contaminación", style);
        getClasificacionParams(row, hyqiData.getContaminacion(), style);
        createCell(row, 3, hyqiData.getContaminacion(), style);   

        row = sheet1.createRow(32);
        createCell(row, 1,"HYQI", style);   
        createFormulaCell(row,3,"SUM(D18:D31)",style); //total sum
        // hyqi from hyqiData
        int hyqi = hyqiData.getAgriculturaDer() + hyqiData.getAgriculturaIzq() + 
          hyqiData.getCanal() + hyqiData.getCoberturaDer() + hyqiData.getCoberturaIzq() +
          hyqiData.getContaminacion() + hyqiData.getEfluentes() + hyqiData.getEstabilidad() + 
          hyqiData.getEstado() + hyqiData.getHumano() + hyqiData.getPresas() + 
          hyqiData.getRegimenes() + hyqiData.getSustrato() + hyqiData.getUrbano();  
        if (hyqi >= 85)
            createCell(row, 2,"Óptimo", style); 
        else if (hyqi >= 47 )
            createCell(row, 2,"Medio", style); 
        else if (hyqi >= 13 )
            createCell(row, 2,"Malo", style);
        else
            createCell(row, 2,"Pobre", style);

        SheetConditionalFormatting sheetCF = sheet1.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1  = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, "85");
        // Condition 1: Cell Value Is   greater than 85   (Blue Fill)
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 2: Cell Value Is  greater than 47  (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, "47");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.LIME.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 3: Cell Value Is   greater than 13   (yellow Fill)
        ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, "13");
        PatternFormatting fill3 = rule3.createPatternFormatting();
        fill3.setFillBackgroundColor(IndexedColors.YELLOW.index);
        fill3.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 4: Cell Value Is  less than 12  (red Fill)
        ConditionalFormattingRule rule4 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "12");
        PatternFormatting fill4 = rule4.createPatternFormatting();
        fill4.setFillBackgroundColor(IndexedColors.RED.index);
        fill4.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        CellRangeAddress[] regions = { CellRangeAddress.valueOf("D33") };
  
        sheetCF.addConditionalFormatting(regions, rule1);
        sheetCF.addConditionalFormatting(regions, rule2);
        sheetCF.addConditionalFormatting(regions, rule3);
        sheetCF.addConditionalFormatting(regions, rule4);
        sheet1.autoSizeColumn(1);
        sheet1.autoSizeColumn(4);
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
        font.setFontHeight(10);
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

    private void getClasificacionParams(Row row, int puntaje, CellStyle style){
        switch (puntaje) {
            case 1:
                createCell(row, 2, "Pobre", style);
                break;
            case 4:
                createCell(row, 2, "Malo", style);
                break;
            case 7:
                createCell(row, 2, "Medio", style);
                break;
            case 10:
                createCell(row, 2, "Óptimo", style);
                break;
            default:
                createCell(row, 2, "ERROR", style);
                break;
        }
    }

    private void getClasificacionParamsRiberas(Row row, int puntaje, CellStyle style){
        switch (puntaje) {
            case 0:
                createCell(row, 2, "Pobre", style);
                break;
            case 2:
                createCell(row, 2, "Malo", style);
                break;
            case 3:
                createCell(row, 2, "Medio", style);
                break;
            case 5:
                createCell(row, 2, "Óptimo", style);
                break;
            default:
                createCell(row, 2, "ERROR", style);
                break;
        }
    }

    private void writeProtocolRef() {   
        // merging cells for sheet title 
        sheet2.addMergedRegion(new CellRangeAddress(1,1, 0, 4));
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(9);
        style.setFont(font);  // style for header title

        Row row = sheet2.createRow(1);
        createBoldCell(row,0, "I. CUENCA");  
        row = sheet2.createRow(2);
        createBoldCell(row,0, "PARÁMETRO");   
        createBoldCell(row,1, "ÓPTIMO");   
        createBoldCell(row,2, "MEDIO");   
        createBoldCell(row,3, "MALO");   
        createBoldCell(row,4, "POBRE");
        row = sheet2.createRow(3);
        createCell(row,0, "1. Cobertura vegetal", style);   
        createCell(row,1, "Más de 70% de cobertura vegetal con especies nativas",
            style);   
        createCell(row,2, "60-40% de cobertura vegetaL nativa", style);   
        createCell(row,3, "50-30% de la ribera cubierta por vegetación nativa", 
            style);   
        createCell(row,4, "Menos de 30 % de superficie cubierta por vegetación",
            style);   
        row = sheet2.createRow(4);
        createCell(row,0, "Ribera derecha", style);   
        createCell(row,1, "5", style);   
        createCell(row,2, "3", style);   
        createCell(row,3, "2", style);   
        createCell(row,4, "0", style);
        row = sheet2.createRow(5);
        createCell(row,0, "Ribera izquierda", style);   
        createCell(row,1, "5", style);   
        createCell(row,2, "3", style);   
        createCell(row,3, "2", style);   
        createCell(row,4, "0", style);
        row = sheet2.createRow(6);
        createCell(row,0, "2. Estabilidad del banco", style);   
        createCell(row,1, "Poca o mínima (<10%) evidencia de erosión", style);   
        createCell(row,2, "Pequeñas áreas de erosión (10-50 %)", style);   
        createCell(row,3, "Potencial de erosión (50-80%) durante inundaciones", 
            style);   
        createCell(row,4, "Muchas áreas erosionadas (>80%)", style);
        row = sheet2.createRow(7);
        writeScoreRef(row, style);
        row = sheet2.createRow(8);
        createCell(row,0, "3. Características del sustrato.", style);   
        createCell(row,1, "Grava, arena; raíces sumergidas, vegetación acuática", 
            style);   
        createCell(row,2, "Arena,arcilla o lodo; algunas raíces sumergidas y vegetación", style);   
        createCell(row,3, "Arcilla en superficie, pocas raíces, sin vegetación", 
            style);   
        createCell(row,4, "Capa de arcilla o rocas, sin raíces o vegetación", 
            style);
        row = sheet2.createRow(9);
        writeScoreRef(row, style);
        row = sheet2.createRow(10);
        createCell(row,0, "4. Desarrollo ganadería y agricultura en la ribera",
            style);   
        createCell(row,1, "Sin cultivos o zonas para ganado", style);   
        createCell(row,2, "20% del suelo para uso agrícola y ganadero", style);   
        createCell(row,3, "50% del suelo para uso agrícola y ganadero", style);   
        createCell(row,4, "Más del 80% del suelo para estos usos ", 
            style);
        row = sheet2.createRow(11);
        createCell(row,0, "Ribera derecha", style);   
        createCenteredCell(row,1, "5", style);   
        createCenteredCell(row,2, "3", style);   
        createCenteredCell(row,3, "2", style);   
        createCenteredCell(row,4, "0", style);
        row = sheet2.createRow(12);
        createCenteredCell(row,0, "Ribera izquierda", style);   
        createCenteredCell(row,1, "5", style);   
        createCenteredCell(row,2, "3", style);   
        createCenteredCell(row,3, "2", style);   
        createCenteredCell(row,4, "0", style);

        sheet2.addMergedRegion(new CellRangeAddress(13,13, 0, 4));
        row = sheet2.createRow(13);
        createBoldCell(row,0, "II. HIDROLOGÍA");  
        row = sheet2.createRow(14);
        createBoldCell(row,0, "PARÁMETRO");   
        createBoldCell(row,1, "ÓPTIMO");   
        createBoldCell(row,2, "MEDIO");   
        createBoldCell(row,3, "MALO");   
        createBoldCell(row,4, "POBRE");
        row = sheet2.createRow(15);
        sheet2.addMergedRegion(new CellRangeAddress(15,15, 1, 2));
        sheet2.addMergedRegion(new CellRangeAddress(15,15, 3, 4));
        createCell(row,0, "5. Presencia de presas", style);   
        createCell(row,1, "Ausencia de presas (incluyendo de gavión y de costales)", style);   
        createCell(row,3, "Presencia de presas (incluyendo de gavión y de costales)", style);   
        row = sheet2.createRow(16);
        sheet2.addMergedRegion(new CellRangeAddress(16,16, 1, 2));
        sheet2.addMergedRegion(new CellRangeAddress(16,16, 3, 4));
        createCell(row,0, "Puntaje", style);   
        createCell(row,1, "10", style);   
        createCell(row,3, "0", style);  
        row = sheet2.createRow(17);
        createCell(row,0, "6. Regímenes de velocidad/profundidad", style);   
        createCell(row,1, "4 reg: lento-profundo,lento-somero,rápido-profundo,rápido-somero",
            style);   
        createCell(row,2, "3 regímenes", style);   
        createCell(row,3, "2 regímenes", style);   
        createCell(row,4, "1 régimen (usualmente lento-somero)", style); 
        row = sheet2.createRow(18);
        writeScoreRef(row, style);
        row = sheet2.createRow(19);
        createCell(row,0, "7. Alteración en el canal", style);   
        createCell(row,1, "Ausencia de canalización", style);   
        createCell(row,2, "Evidencia de canalización en el pasado", style);   
        createCell(row,3, "Canalización extensiva, 40-80% canalizado e interrumpido", style);   
        createCell(row,4, "Banco con cemento o gavión, +80% canalizado", style); 
        row = sheet2.createRow(20);
        writeScoreRef(row, style);
        row = sheet2.createRow(21);
        createCell(row,0, "8. Estado del canal", style);   
        createCell(row,1, "Agua hasta la base de ambos bancos,sustrato expuesto mínimamente", style);   
        createCell(row,2, "Agua llena >75% del canal,ó 25% del sustrato está expuesto", style);   
        createCell(row,3, "Agua llena 25-75% del canal o el sustrato está expuesto", style);   
        createCell(row,4, "Muy poca agua en el canal", style); 
        row = sheet2.createRow(22);
        writeScoreRef(row, style);

        sheet2.addMergedRegion(new CellRangeAddress(23,23, 0, 4));
        row = sheet2.createRow(23);
        createBoldCell(row,0, "III. PERTURBACIONES ANTROPOGÉNICAS");  
        row = sheet2.createRow(24);
        createBoldCell(row,0, "PARÁMETRO");   
        createBoldCell(row,1, "ÓPTIMO");   
        createBoldCell(row,2, "MEDIO");   
        createBoldCell(row,3, "MALO");   
        createBoldCell(row,4, "POBRE");
        row = sheet2.createRow(25);
        sheet2.addMergedRegion(new CellRangeAddress(25,25, 1, 2));
        sheet2.addMergedRegion(new CellRangeAddress(25,25, 3, 4));
        createCell(row,0, "9. Efluentes directos al río por el uso doméstico", 
            style);   
        createCell(row,1, "Ausencia", style);   
        createCell(row,3, "Presencia", style);   
        row = sheet2.createRow(26);
        sheet2.addMergedRegion(new CellRangeAddress(26,26, 1, 2));
        sheet2.addMergedRegion(new CellRangeAddress(26,26, 3, 4));
        createCell(row,0, "Puntaje", style);   
        createCell(row,1, "10", style);   
        createCell(row,3, "1", style);  
        row = sheet2.createRow(27);
        createCell(row,0, "10. Desarrollo urbano", style);   
        createCell(row,1, "Sin asentamientos,carreteras, derivaciones p/usos domésticos o industriales", style);   
        createCell(row,2, "20% del suelo para uso humano", style);   
        createCell(row,3, "50% del suelo para uso humano", style);   
        createCell(row,4, "Más del 80% del suelo para uso humano", style); 
        row = sheet2.createRow(28);
        writeScoreRef(row, style);
        row = sheet2.createRow(29);
        createCell(row,0, "11. Desarrollo humano", style);   
        createCell(row,1, "Ausencia de actividades humanas", style);   
        createCell(row,2, "Al menos una actividad: ganadera, agrícola o piscícola",
             style);   
        createCell(row,3, "Al menos 3 actividades: agrícola, ganadera, piscícola, doméstica", 
            style);   
        createCell(row,4, "Actividades agrícolas, ganaderas, piscícolas y domésticas", 
            style); 
        row = sheet2.createRow(30);
        writeScoreRef(row, style);
        row = sheet2.createRow(31);
        createCell(row,0, "12. Presencia de contaminación,basura y escombros", 
            style);   
        createCell(row,1, "Menos del 10% de presencia de basura y/o escombros", 
            style);   
        createCell(row,2, "Entre 20-40 % de presencia de basura y/o escombros", 
            style);   
        createCell(row,3, "Entre 50-80 % de presencia de basura y/o escombros", 
            style); 
        createCell(row,4, "Más del 90% de basura y/o escombros", style); 
        row = sheet2.createRow(32);
        writeScoreRef(row, style);

        sheet2.autoSizeColumn(0);
        sheet2.autoSizeColumn(1);
        sheet2.autoSizeColumn(2);
        sheet2.autoSizeColumn(3);
        sheet2.autoSizeColumn(4);
    }

    private void writeScoreRef(Row row, CellStyle style){
        createCenteredCell(row,0, "Puntaje", style);   
        createCenteredCell(row,1, "10", style);   
        createCenteredCell(row,2, "7", style);   
        createCenteredCell(row,3, "4", style);   
        createCenteredCell(row,4, "1", style);
    }
       
    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeAuthorLines();
        writeDataHeaderLine();
        writeDataLines();
        writeProtocolRef();
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close();
         
        outputStream.flush();
        outputStream.close();    
    }

}