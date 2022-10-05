package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;


import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.model.DeqiData;
import com.ciencias.CUME.api.service.DeqiService;
 
@RestController
@RequestMapping("/api/deqi")
public class DeqiController {
 

    @PostMapping(path="/export/excel")
    public void exportToExcel(HttpServletResponse response, @RequestBody DeqiData data) throws IOException {
        
        response.reset();
        System.out.println(data.getAchnanthesCoarctata());

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=CalidadHidromorfologica.xlsx");

        DeqiService deqiService = new DeqiService(data);
        deqiService.export(response);
          
    } 
} 
 
