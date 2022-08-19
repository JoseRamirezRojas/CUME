package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;


import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.model.HyqiData;
import com.ciencias.CUME.api.service.HyqiService;
 
@RestController
@RequestMapping("/api/hyqi")
public class HyqiController {
 

    @PostMapping(path="/export/excel")
    public void exportToExcel(HttpServletResponse response, @RequestBody HyqiData data) throws IOException {
        
        response.reset();
        System.out.println(data.getCoberturaDer());

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=CalidadHidromorfologica.xlsx");

        HyqiService physChemService = new HyqiService(data);
        physChemService.export(response);
          
    } 
} 
 
