package com.ciencias.CUME.api.service;

import java.io.IOException;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.DataFormat;
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
        sheet1.addMergedRegion(new CellRangeAddress(0,0, 0, 3));
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
        createCell(row,0, "Código de la muestra:", style);
        createCell(row,2, "DEQI:", style);
        createFormulaCell(row,3, "G11", style);
        row = sheet1.createRow(6);   
        createCell(row,0, "Dilución:", style);
        createCell(row,2, "Calidad ecológica:", style);
        createFormulaCell(row,3, "H11", style);  
        
        sheet1.autoSizeColumn(0);
        sheet1.autoSizeColumn(2);
        sheet1.autoSizeColumn(3);
    }

    private void writeDataHeaderLine() {
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);

        Row row = sheet1.createRow(9);   
        createBoldCell(row,2, "Valvas totales");
        createBoldItalicCell(row,3, "Σ h");
        createBoldItalicCell(row,4, "Σ h · v");
        createBoldCell(row,6, "DEQI");
        createBoldCell(row,7, "Calidad");

        row = sheet1.createRow(10); 
        createFormulaCell(row,2,"SUM(C14:C500)",style);
        createFormulaCell(row,3,"SUM(D14:D500)",style);
        createFormulaCell(row,4,"SUM(E14:E500)",style);
        createFormulaCell(row,6,"E11/D11",style);

        row = sheet1.createRow(11);        
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 0, 0));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 1, 1));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 2, 2));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 3, 3));
        sheet1.addMergedRegion(new CellRangeAddress(11,12, 4, 4));
        createItalicCell(row, 0, "Valor indicador (v)");
        createCenteredCell(row, 1, "Especie", style);
        createCenteredCell(row, 2, "Conteo", style);
        createItalicCell(row, 3, "Abundancia (h)");
        createItalicCell(row, 4, "h·v");
        
        sheet1.autoSizeColumn(3,true);
        sheet1.autoSizeColumn(4,true);
        
        int numMerged = sheet1.getNumMergedRegions();
        for(int i = 1; i<numMerged;i++){  // apply border to merged cells
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

        int rownum = 13;
        Row row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Achnanthes coarctata (Brébisson ex Kützing) Grunow", style);
        createCenteredCell(row,2, deqiData.getAchnanthesCoarctata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Achnanthidium exiguum (Grunow) Czarnecki", style);
        createCenteredCell(row,2, deqiData.getAchnanthidiumExiguum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Achnanthidium minutissimum (Kützing) Czarnecki", style);
        createCenteredCell(row,2, deqiData.getAchnanthidiumMinutissimum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Adlafia minuscula (Grunow) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getAdlafiaMinuscula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Amphora copulata (Kützing) Schoeman & R.E.M. Archibald", style);
        createCenteredCell(row,2, deqiData.getAmphoraCopulata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Amphora pediculus (Kützing) Grunow ex A. Schmidt", style);
        createCenteredCell(row,2, deqiData.getAmphoraPediculus(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Aulacoseira ambigua (Grunow) Simonsen", style);
        createCenteredCell(row,2, deqiData.getAulacoseiraAmbigua(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Caloneis bacillum (Grunow) Cleve", style);
        createCenteredCell(row,2, deqiData.getCaloneisBacillum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Caloneis fontinalis (Grunow) Lange-Bertalot & Reichardt", style);
        createCenteredCell(row,2, deqiData.getCaloneisFontinalis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Caloneis silicula (Ehrenberg) Cleve", style);
        createCenteredCell(row,2, deqiData.getCaloneisSilicula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Caloneis stauroneiformis (Amossé) Metzeltin & Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getCaloneisStauroneiformis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Cavinula cocconeiformis (Gregory ex Greville) D.G. Mann & A.J. Stickle", style);
        createCenteredCell(row,2, deqiData.getCavinulaCocconeiformis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Cavinula lapidosa (Krasske) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getCavinulaLapidosa(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Cavinula pseudoscutiformis (Hustedt) D.G. Mann & A.J. Stickle", style);
        createCenteredCell(row,2, deqiData.getCavinulaPseudoscutiformis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Chamaepinnularia submuscicola (Krasske) Lange-Bertalot in Moser et al.", style);
        createCenteredCell(row,2, deqiData.getChamaepinnulariaSubmuscicola(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Cocconeis placentula (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getCocconeisPlacentula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Craticula subminuscula (Manguin) Wetzel & Ector", style);
        createCenteredCell(row,2, deqiData.getCraticulaSubminuscula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Cyclotella menenghiniana (Kützing)", style);
        createCenteredCell(row,2, deqiData.getCyclotellaMenenghiniana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Cymbella mexicana (Ehrenberg) Cleve", style);
        createCenteredCell(row,2, deqiData.getCymbellaMexicana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Cymbella tumida (Brébisson) Van Heurk", style);
        createCenteredCell(row,2, deqiData.getCymbellaTumida(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Cymbopleura naviculiformis (Auerswald ex Heiberg) Krammer", style);
        createCenteredCell(row,2, deqiData.getCymbopleuraNaviculiformis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Decussata placenta (Ehrenberg) Lange-Bertalot & Metzeltin", style);
        createCenteredCell(row,2, deqiData.getDecussataPlacenta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Diploneis smithii (Brébisson) Cleve", style);
        createCenteredCell(row,2, deqiData.getDiploneisSmithii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Diploneis subovalis (Cleve)", style);
        createCenteredCell(row,2, deqiData.getDiploneisSubovalis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Discostella pseudostelligera (Hustedt) Hout & Klee", style);
        createCenteredCell(row,2, deqiData.getDiscostellaPseudostelligera(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Encyonema lange-bertalotii (Krammer)", style);
        createCenteredCell(row,2, deqiData.getEncyonemaLangeBertalotii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Encyonema minutum (Hilse ex Rabenhorst) Mann", style);
        createCenteredCell(row,2, deqiData.getEncyonemaMinutum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Encyonema silesiacum (Blesich) D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getEncyonemaSilesiacum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Encyonema ventricosum (Kützing) Grunow in Schmidt et al.", style);
        createCenteredCell(row,2, deqiData.getEncyonemaVentricosum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Epithemia adnata (Kützing) Brébisson", style);
        createCenteredCell(row,2, deqiData.getEpithemiaAdnata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Epithemia turgida (Ehrenberg) Kützing", style);
        createCenteredCell(row,2, deqiData.getEpithemiaTurgida(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Eunotia arcus (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getEunotiaArcus(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Eunotia bilunaris (Ehrenberg) Schaarschmidt", style);
        createCenteredCell(row,2, deqiData.getEunotiaBilunaris(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Eunotia implicata (Nörpel, Lange-Bertalot & Alles)", style);
        createCenteredCell(row,2, deqiData.getEunotiaImplicata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Eunotia minor (Kützing) Grunow", style);
        createCenteredCell(row,2, deqiData.getEunotiaMinor(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Eunotia paratridentula (Lange-Bertalot & Kulikovskiy)", style);
        createCenteredCell(row,2, deqiData.getEunotiaParatridentula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Fragilaria capucina var capucina (Desmazières)", style);
        createCenteredCell(row,2, deqiData.getFragilariaCapucina(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Fragilaria crotonensis (Kitton)", style);
        createCenteredCell(row,2, deqiData.getFragilariaCrotonensis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Fragilaria vaucheriae (Kützing) J.B. Petersen", style);
        createCenteredCell(row,2, deqiData.getFragilariaVaucheriae(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Frankophila sp.", style);
        createCenteredCell(row,2, deqiData.getFrankophila(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Frustulia crassinervia (Brébisson) Lange-Bertalot & Krammer", style);
        createCenteredCell(row,2, deqiData.getFrustuliaCrassinervia(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Frustulia vulgaris (Thwaites) De Toni", style);
        createCenteredCell(row,2, deqiData.getFrustuliaVulgaris(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Geissleria acceptata (Hustedt) Lange-Bertalot & Metzeltin", style);
        createCenteredCell(row,2, deqiData.getGeissleriaAcceptata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Gomphonema acuminatum (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaAcuminatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Gomphonema capitatum (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaCapitatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Gomphonema clavatum (Krammer & Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaClavatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Gomphonema commutatum (Grunow)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaCommutatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Gomphonema gracile (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaGracile(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Gomphonema lagenula (Kützing)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaLagenula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Gomphonema minutum f. pachypus (Lange-Bertalot & Reichardt)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaMinutum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Gomphonema parvulum (Kützing)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaParvulum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Gomphonema tenuissimum (Fricke in Schmidt et al.)", style);
        createCenteredCell(row,2, deqiData.getGomphonemaTenuissimum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Halamphora montana (Krasske) Levkov", style);
        createCenteredCell(row,2, deqiData.getHalamphoraMontana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Halamphora veneta (Kützing) Levkov", style);
        createCenteredCell(row,2, deqiData.getHalamphoraVeneta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Hantzschia amphioxys (Ehrenberg) Grunow", style);
        createCenteredCell(row,2, deqiData.getHantzschiaAmphioxys(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Hantzschia calcifuga (Reichardt & Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getHantzschiaCalcifuga(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Humidophila contenta (Grunow) Lowe et al.", style);
        createCenteredCell(row,2, deqiData.getHumidophilaContenta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Humidophila perpusilla (Grunow) Lowe et al.", style);
        createCenteredCell(row,2, deqiData.getHumidophilaPerpusilla(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Lemnicola hungarica (Grunow) T.E. Round & P.W. Basson", style);
        createCenteredCell(row,2, deqiData.getLemnicolaHungarica(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Luticola goeppertiana (Bleisch) D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getLuticolaGoeppertiana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Luticola mutica (Kützing) D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getLuticolaMutica(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Luticola nivalis (Ehrenberg) D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getLuticolaNivalis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Mayamaea atomus (Kützing) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getMayamaeaAtomus(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Melosira varians (C. Agardh)", style);
        createCenteredCell(row,2, deqiData.getMelosiraVarians(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Meridon constrictum (Ralfs)", style);
        createCenteredCell(row,2, deqiData.getMeridonConstrictum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula angusta (Grunow)", style);
        createCenteredCell(row,2, deqiData.getNaviculaAngusta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula cryptocephala (Kützing)", style);
        createCenteredCell(row,2, deqiData.getNaviculaCryptocephala(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula cryptotenella (Kützing)", style);
        createCenteredCell(row,2, deqiData.getNaviculaCryptotenella(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula gregaria (Donkin)", style);
        createCenteredCell(row,2, deqiData.getNaviculaGregaria(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula radiosa (Kützing)", style);
        createCenteredCell(row,2, deqiData.getNaviculaRadiosa(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula rhynocephala (Kützing)", style);
        createCenteredCell(row,2, deqiData.getNaviculaRhynchocephala(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Navicula seibigiana (Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getNaviculaSeibigiana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Navicula symmetrica (Patrick)", style);
        createCenteredCell(row,2, deqiData.getNaviculaSymmetrica(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Navicula tenelloides (Hustedt)", style);
        createCenteredCell(row,2, deqiData.getNaviculaTenelloides(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Navicula veneta Kützing", style);
        createCenteredCell(row,2, deqiData.getNaviculaVeneta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Navicula vilaplanii (Lange-Bertalot & Sabater)", style);
        createCenteredCell(row,2, deqiData.getNaviculaVilaplanii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Neidium ampliatum (Ehrenberg) Krammer ", style);
        createCenteredCell(row,2, deqiData.getNeidiumAmpliatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Neidium sacoense (Reimer)", style);
        createCenteredCell(row,2, deqiData.getNeidiumSacoense(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Neidium sacoense (Reimer)", style);
        createCenteredCell(row,2, deqiData.getNeidiumSacoense(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia acidoclinata (Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaAcidoclinata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Nitzschia acula (Kützing) Hantzsch", style);
        createCenteredCell(row,2, deqiData.getNitzschiaAcula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia bacillariaeformis (Hustedt)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaBacillariaeformis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia claussi (Hantzsch)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaClaussi(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia communis (Rabenhorst)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaCommunis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Nitzschia costei (L. Tudesque, F. Rimet & L. Ector)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaCostei(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia dissipata (Kützing) Grunow", style);
        createCenteredCell(row,2, deqiData.getNitzschiaDissipata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia fonticola (Grunow)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaFonticola(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Nitzschia frustulum (Kützing) Grunow", style);
        createCenteredCell(row,2, deqiData.getNitzschiaFrustulum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia linearis (W. Smith)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaLinearis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Nitzschia palea (Kützing) W. Smith", style);
        createCenteredCell(row,2, deqiData.getNitzschiaPalea(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia paleacea (Grunow) Grunow in Van Heurck", style);
        createCenteredCell(row,2, deqiData.getNitzschiaPaleacea(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia pusilla (Kützing) Grunow", style);
        createCenteredCell(row,2, deqiData.getNitzschiaPusilla(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Nitzschia recta (Hantzsch ex Rabenhorst)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaRecta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia solgensis (Cleve-Euler)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaSolgensis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia soratensis (E.A. Morales & M.L. Vis)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaSoratensis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Nitzschia supralitorea (Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getNitzschiaSupralitorea(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Nitzschia umbonata (Ehrenberg) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getNitzschiaUmbonata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Nupela praecipua (Reichardt) ", style);
        createCenteredCell(row,2, deqiData.getNupelaPraecipua(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Odontidium mesodon (Kützing) ", style);
        createCenteredCell(row,2, deqiData.getOdontidiumMesodon(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Odontidium rostratum (Z. Levkov & I. Jüttner)", style);
        createCenteredCell(row,2, deqiData.getOdontidiumRostratum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Orthoseira roeseana (Rabenhorst) O'Meara ", style);
        createCenteredCell(row,2, deqiData.getOrthoseiraRoeseana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia anglica (Krammer) ", style);
        createCenteredCell(row,2, deqiData.getPinnulariaAnglica(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia appendiculata var. amaniana (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaAppendiculata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia bertrandii (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaBertrandii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia borealis (Ehrenbeg)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaBorealis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia divergens (W. Smith) ", style);
        createCenteredCell(row,2, deqiData.getPinnulariaDivergens(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia divergentissima (Grunow) Cleve ", style);
        createCenteredCell(row,2, deqiData.getPinnulariaDivergentissima(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia erratica (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaErratica(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia frequentis (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaFrequentis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia inconstans (A.Mayer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaInconstans(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia johndonatoi (Metzeltin & Lange-Bertalot)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaJohndonatoi(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia microstauron (Ehrenberg) Cleve", style);
        createCenteredCell(row,2, deqiData.getPinnulariaMicrostauron(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia nodosa (Ehrenberg) W. Smith ", style);
        createCenteredCell(row,2, deqiData.getPinnulariaNodosa(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Pinnularia peragalli (Kulikovskiy,Lange-Bertalot & Metzeltin)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaNodosa(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Pinnularia reichardtii (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaReichardtii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Pinnularia sinistra (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaSinistra(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia subcommutata var. nonfasciata (Krammer)", style);
        createCenteredCell(row,2, deqiData.getPinnulariaSubcommutata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pinnularia sp.", style);
        createCenteredCell(row,2, deqiData.getPinnularia(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Planothidium biporomum (Hohn & Hellerman) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getPlanothidiumBiporomum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Planothidium frequentissimum (Lange-Bertalot) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getPlanothidiumFrequentissimum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Planothidium lanceolatum (Brébisson ex Kützing) Bukhtiyarova", style);
        createCenteredCell(row,2, deqiData.getPlanothidiumLanceolatum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Platessa conspicua (A. Mayer) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getPlatessaConspicua(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Prestauroneis protractoides (Hustedt) Liu & Kociolek", style);
        createCenteredCell(row,2, deqiData.getPrestauroneisProtractoides(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Pseudostaurosira brevistriata (Grunow in Van Heurk) Williams & Round", style);
        createCenteredCell(row,2, deqiData.getPseudostaurosiraBrevistriata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Pseudostaurosira margaritae (Salinas, Mora, Abarca & Jahn)", style);
        createCenteredCell(row,2, deqiData.getPseudostaurosiraMargaritae(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Reimeira sinuata (Gregory) Kociolek & Stoermer", style);
        createCenteredCell(row,2, deqiData.getReimeiraSinuata(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Rhoicosphenia sp.", style);
        createCenteredCell(row,2, deqiData.getRhoicosphenia(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Rhopalodia rupestris (W. Smith) Krammer", style);
        createCenteredCell(row,2, deqiData.getRhopalodiaRupestris(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Rossithidium nodosum  (Cleve) Aboal", style);
        createCenteredCell(row,2, deqiData.getRossithidiumNodosum(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Sellaphora cosmopolitana (Lange-Bertalot) Wetzel & Ector", style);
        createCenteredCell(row,2, deqiData.getSellaphoraCosmopolitana(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Sellaphora laevissima (Kützing) D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getSellaphoraLaevissima(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Sellaphora nigri (De Not.) C.E. Wetzel & Ector", style);
        createCenteredCell(row,2, deqiData.getSellaphoraNigri(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Sellaphora pseudopupula (Krasske) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getSellaphoraPseudopupula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 5, "Sellaphora pupula (Kützing) Mereschkowsky", style);
        createCenteredCell(row,2, deqiData.getSellaphoraPupula(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Sellaphora saugerresii (Desmazières) C.E. Wetzel & D.G. Mann", style);
        createCenteredCell(row,2, deqiData.getSellaphoraSaugerresii(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Sellaphora sp.", style);
        createCenteredCell(row,2, deqiData.getSellaphora(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Simonsenia delognei (Grunow ex Van Heurk) Lange-Bertalot", style);
        createCenteredCell(row,2, deqiData.getSimonseniaDelognei(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Stauroneis kriegeri (Patrick)", style);
        createCenteredCell(row,2, deqiData.getStauroneisKriegeri(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Stauroneis phoenicentron (Nitzsch) Ehrenberg", style);
        createCenteredCell(row,2, deqiData.getStauroneisPhoenicentron(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Stauroneis subgracilis (Lange-Bertalot & Krammer)", style);
        createCenteredCell(row,2, deqiData.getStauroneisSubgracilis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Stauroneis thermicola (J.B. Petersen) Lund", style);
        createCenteredCell(row,2, deqiData.getStauroneisThermicola(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Staurosira venter (Ehrenberg) Cleve & Möller", style);
        createCenteredCell(row,2, deqiData.getStaurosiraVenter(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Staurosirella leptostauron var. dubia (Grunow) M.B. Edlund", style);
        createCenteredCell(row,2, deqiData.getStaurosirellaLeptostauron(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Stephanodiscus niagarae (Ehrenberg)", style);
        createCenteredCell(row,2, deqiData.getStephanodiscusNiagarae(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 1, "Stephanodiscus oregonicus (Ralfs) Håkansson", style);
        createCenteredCell(row,2, deqiData.getStephanodiscusOregonicus(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Surirella angusta (Kützing)", style);
        createCenteredCell(row,2, deqiData.getSurirellaAngusta(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Surirella linearis (W. Smith)", style);
        createCenteredCell(row,2, deqiData.getSurirellaLinearis(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 3, "Surirella muscicola (Krasske) Lange-Bertalot & Metzeltin", style);
        createCenteredCell(row,2, deqiData.getSurirellaMuscicola(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 4, "Ulnaria acus (Kützing) M. Aboal", style);
        createCenteredCell(row,2, deqiData.getUlnariaAcus(), style);
        row = sheet1.createRow(rownum++);
        insertDeqiDataRow(row, rownum, 2, "Ulnaria ulna (Nitzch) P. Compère", style);
        createCenteredCell(row,2, deqiData.getUlnariaUlna(), style);      
        sheet1.autoSizeColumn(1);
        sheet1.autoSizeColumn(4);
    }

    public void insertDeqiDataRow (Row row,int i, int v, String taxa, CellStyle style){
        createCell(row, 0, v, style);
        createTaxaCell(row, 1, taxa);
        createFormulaCell(row, 3, "(C"+(i)+"/C11)*100", style);
        createFormulaCell(row, 4, "A"+(i)+"*D"+(i), style);

        sheet1.autoSizeColumn(4);
    }

    public void insertDeqiQuality(){
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        style.setFont(font);
        // deqi from deqiData
        CellReference cellReference = new CellReference("G11");
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator(); 
        Row row = sheet1.getRow(cellReference.getRow());
        Cell cell = row.getCell(cellReference.getCol()); 
        CellValue cellValue = evaluator.evaluate(cell);     
        switch (cellValue.getCellType()) {
            case BOOLEAN:
                break;
            case NUMERIC:
                double deqi = cellValue.getNumberValue();
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
                break;
            default:
                break;
        }

        // conditional formatting of DEQI
        SheetConditionalFormatting sheetCF = sheet1.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1  = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "1.5");
        // Condition 1: Cell Value Is LESSS than 1.5   (Blue Fill)
        PatternFormatting fill1 = rule1.createPatternFormatting();
        fill1.setFillBackgroundColor(IndexedColors.LIGHT_BLUE.index);
        fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 2: Cell Value Is  LESS than 2.5  (Green Fill)
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "2.5");
        PatternFormatting fill2 = rule2.createPatternFormatting();
        fill2.setFillBackgroundColor(IndexedColors.LIME.index);
        fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 3: Cell Value Is   LESS than 3.5   (yellow Fill)
        ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "3.5");
        PatternFormatting fill3 = rule3.createPatternFormatting();
        fill3.setFillBackgroundColor(IndexedColors.YELLOW.index);
        fill3.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 4: Cell Value Is  LESS than 4.5  (red Fill)
        ConditionalFormattingRule rule4 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "4.5");
        PatternFormatting fill4 = rule4.createPatternFormatting();
        fill4.setFillBackgroundColor(IndexedColors.RED.index);
        fill4.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        // Condition 5: Cell Value Is  LESS than 5  (dark red Fill)
        ConditionalFormattingRule rule5 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LE, "5");
        PatternFormatting fill5 = rule5.createPatternFormatting();
        fill5.setFillBackgroundColor(IndexedColors.DARK_RED.index);
        fill5.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
        CellRangeAddress[] regions = { CellRangeAddress.valueOf("G11") };
        sheetCF.addConditionalFormatting(regions, rule1);
        sheetCF.addConditionalFormatting(regions, rule2);
        sheetCF.addConditionalFormatting(regions, rule3);
        sheetCF.addConditionalFormatting(regions, rule4);
        sheetCF.addConditionalFormatting(regions, rule5);

        sheet1.autoSizeColumn(6);
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

    private void createBoldItalicCell(Row row, int columnCount, Object value ) {
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

    private void createItalicCell(Row row, int columnCount, Object value ) {
        Cell cell = row.createCell(columnCount);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
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
        sheet1.autoSizeColumn(columnCount);
    }

    private void createTaxaCell(Row row, int columnCount, Object value ) {
        Cell cell = row.createCell(columnCount);
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(10);
        font.setItalic(true);
        style.setFont(font);
        cell.setCellValue((String) value);
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
        CellStyle formulaStyle = cell.getCellStyle();
        formulaStyle.setDataFormat(format.getFormat("0.00"));
        XSSFFont font = workbook.createFont();
        font.setFontHeight(12);
        formulaStyle.setFont(font);
        formulaStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellFormula(formula);
        cell.setCellStyle(formulaStyle);
    }

    private void writeCollectSheet(){
        sheet2.addMergedRegion(new CellRangeAddress(1,1, 0, 5));
        sheet2.addMergedRegion(new CellRangeAddress(2,2, 1, 3));
        sheet2.addMergedRegion(new CellRangeAddress(4,4, 0, 5));
        sheet2.addMergedRegion(new CellRangeAddress(6,6,0, 5));
        sheet2.addMergedRegion(new CellRangeAddress(7,7, 0, 1));
        sheet2.addMergedRegion(new CellRangeAddress(7,7, 2, 3));
        sheet2.addMergedRegion(new CellRangeAddress(7,7, 4, 5));
        sheet2.addMergedRegion(new CellRangeAddress(8,8, 0, 1));
        sheet2.addMergedRegion(new CellRangeAddress(8,8, 2, 3));
        sheet2.addMergedRegion(new CellRangeAddress(8,8, 4, 5));
        sheet2.addMergedRegion(new CellRangeAddress(9,9, 1, 2));
        sheet2.addMergedRegion(new CellRangeAddress(9,9, 4, 5));
        sheet2.addMergedRegion(new CellRangeAddress(10,10, 1, 5));
        CellStyle style = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setFontHeight(10);
        style.setFont(font);  // style for header title
        int rownum = 1;
        Row row = sheet2.createRow(rownum);
        createBoldCell(row,0, "FICHA DE COLECTA DE DIATOMEAS BENTÓNICAS"); 
        rownum ++;
        row = sheet2.createRow(rownum++);
        createCell(row,0, "Colector", style);
        createCell(row,4, "Código de la muestra", style); 
        row = sheet2.createRow(rownum++);
        createCell(row,0, "Río",style);
        createCell(row,2, "Localidad",style); 
        createCell(row,4, "Fecha",style); 
        row = sheet2.createRow(rownum++);
        createBoldCell(row,0, "Coordenadas");
        row = sheet2.createRow(rownum++);
        createCell(row,0, "Latitud",style);
        createCell(row,2, "Longitud",style); 
        createCell(row,4, "Elevación",style); 
        row = sheet2.createRow(rownum++);
        createBoldCell(row,0, "Sustrato colectado");
        row = sheet2.createRow(rownum++);
        createCenteredCell(row,0, "Cantos/Rocas",style);
        createCenteredCell(row,2, "Artificial",style); 
        createCenteredCell(row,4, "Vegetación",style); 
        rownum ++;
        row = sheet2.createRow(rownum++);
        createCell(row,0, "Número de sustratos colectados:",style);
        createCell(row,3, "Superficie colectada: (cm²)" , style); 
        row = sheet2.createRow(rownum++);
        createCell(row,0, "Preservación:",style);

        CellRangeAddress region = CellRangeAddress.valueOf("A2:F11");
        BorderStyle borderStyle = BorderStyle.MEDIUM ;
        RegionUtil.setBorderBottom(borderStyle, region, sheet2);
        RegionUtil.setBorderTop(borderStyle, region, sheet2);
        RegionUtil.setBorderLeft(borderStyle, region, sheet2);
        RegionUtil.setBorderRight(borderStyle, region, sheet2);

        sheet2.autoSizeColumn(0);
        sheet2.autoSizeColumn(1);
        sheet2.autoSizeColumn(2);
        sheet2.autoSizeColumn(3);
        sheet2.autoSizeColumn(4);
        sheet2.autoSizeColumn(5);
    }

    public void export(HttpServletResponse response) throws IOException {
        writeHeaderLine();
        writeAuthorLines();
        writeDataHeaderLine();
        writeDataLines();
        insertDeqiQuality();
        writeCollectSheet();
           
        ServletOutputStream outputStream = response.getOutputStream();
        
        workbook.write(outputStream);
        workbook.close(); 
        outputStream.flush();
        outputStream.close();    
    }
}