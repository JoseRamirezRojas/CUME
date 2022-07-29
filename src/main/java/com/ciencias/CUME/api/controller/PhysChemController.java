package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;


import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.model.PhysChemData;
import com.ciencias.CUME.api.service.PhysChemService;
 
@RestController
@RequestMapping("/api/physchem")
public class PhysChemController {
 

    @PostMapping(path="/export/excel")
    public void exportToExcel(HttpServletResponse response, @RequestBody PhysChemData data) throws IOException {
        
        response.reset();
        System.out.println(data.getTemperatura3());

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=CalidadFisicoquimica.xlsx");

        PhysChemService physChemService = new PhysChemService(data);
        physChemService.export(response);
          
    } 
} 
 
