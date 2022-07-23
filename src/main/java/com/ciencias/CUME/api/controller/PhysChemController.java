package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.model.PhysChemData;
import com.ciencias.CUME.api.service.PhysChemService;
import com.ciencias.CUME.api.service.UserExcelExporter;
 
@RestController
@RequestMapping("/api/physchem")
public class PhysChemController {
 

    @PostMapping(path="/export/excel")
    public void exportToExcel(HttpServletResponse response, @RequestBody PhysChemData data) throws IOException {
        
        response.reset();

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=CalidadFisicoquimica.xlsx");

         
        //List<User> listUsers = service.listAll();

        PhysChemService physChemService = new PhysChemService(data);
        physChemService.export(response);
         
        // UserExcelExporter excelExporter = new UserExcelExporter();
        
        // excelExporter.export(response);  
    } 
} 
 
