package com.ciencias.CUME.api.controller;

import java.io.IOException;
 
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.HttpHeaders;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.ciencias.CUME.api.service.UserExcelExporter;
 
//@CrossOrigin(origins = "http://localhost:3000") 
@RestController
@RequestMapping("/api/messages")
public class UserController {
 
    private final static Logger LOGGER = LoggerFactory.getLogger(UserController.class);

    // @Autowired
    // private UserServices service;
     
    //@CrossOrigin(origins = "http://localhost:3000") 
    @PostMapping(path="/users/export/excel")
    public void exportToExcel(HttpServletResponse response) throws IOException {
        
        
        
        String headerKey = "Content-Disposition";
        String headerValue = "attachment; filename=usersss.xlsx";

        response.reset();

        response.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=usersss.xlsx");

         
        //List<User> listUsers = service.listAll();
         
        UserExcelExporter excelExporter = new UserExcelExporter();
        
        excelExporter.export(response);  
        
        /**UserExcelExporter excelExporter = new UserExcelExporter();
        response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        response.setHeader("Content-Disposition", "attachment; filename=data.xlsx");
        response.setHeader("Access-Control-Allow-Origin", "*");
        ByteArrayInputStream stream = excelExporter.prepareData();
        IOUtils.copy(stream, response.getOutputStream()); **/
    } 
} 
 
