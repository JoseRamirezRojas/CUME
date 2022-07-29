package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.service.UserExcelExporter;
 
@RestController
@RequestMapping("/api/messages")
public class UserController {
 
    private final static Logger LOGGER = LoggerFactory.getLogger(UserController.class);

    @PostMapping(path="/users/export/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {

        response.reset();

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=usersss.xlsx");

         
        UserExcelExporter excelExporter = new UserExcelExporter();
        
        excelExporter.export(response);  
        
    } 
} 
 
