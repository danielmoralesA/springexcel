package com.example.demo.controller;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.excel.ExcelPoiHelper;
import com.example.demo.excel.MyCell;

import java.util.Collection;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Controller
public class ExcelController {
	
    private String fileLocation;
    
    private ExcelPoiHelper excelPOIHelper;
	
	@RequestMapping(method = RequestMethod.GET, value = "/")
    public String getExcelProcessingPage() {
        return "excel";
    }
	
	@RequestMapping(method = RequestMethod.POST, value = "/upload")
    public String uploadFile(Model model, MultipartFile file) throws IOException {
        InputStream in = file.getInputStream();
        File currDir = new File(".");
        String path = currDir.getAbsolutePath();
        fileLocation = path.substring(0, path.length() - 1) + file.getOriginalFilename();
        FileOutputStream f = new FileOutputStream(fileLocation);
        int ch = 0;
        while ((ch = in.read()) != -1) {
            f.write(ch);
        }
        f.flush();
        f.close();
        //read file
       
        if(fileLocation!=null) {
        	 if (fileLocation.endsWith(".xlsx") || fileLocation.endsWith(".xls")) {
        		 excelPOIHelper = new ExcelPoiHelper();
        		 System.out.println("location "+fileLocation);
        		 Map<Integer, List<MyCell>> data = excelPOIHelper.readExcel(fileLocation);
        		 List<MyCell> values = data.values().stream()
        				 .flatMap(Collection::stream)
                         .collect(Collectors.toList());
        		 for(MyCell cell : values) {
        			 System.out.println("valor "+cell.getContent());
        		 }
                 model.addAttribute("data", data);
             } else {
                 model.addAttribute("message", "Not a valid excel file!");
        	 }
        }
        
        model.addAttribute("message", "File: " + file.getOriginalFilename() + " has been uploaded successfully!");
        return "excel";
    }
	
	

}
