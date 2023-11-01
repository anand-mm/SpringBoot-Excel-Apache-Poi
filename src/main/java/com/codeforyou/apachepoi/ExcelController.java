package com.codeforyou.apachepoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;


@RestController
public class ExcelController {

    @GetMapping("/excel")
    public void excelGenerator() throws Exception{

        FileInputStream templateFile = new FileInputStream("template.xls");
        Workbook workbook = WorkbookFactory.create(templateFile);
        templateFile.close();

        Sheet sheet = workbook.getSheetAt(0); // Assuming you're modifying the first sheet

        String searchString = "Central CPSEs";

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    if (cellValue.contains(searchString)) {
                        System.out.println("Found '" + searchString + "' in cell: " + cell.getAddress());
                    }
                }
            }
        }

        FileOutputStream outputFile = new FileOutputStream("modified_template.xls");
        workbook.write(outputFile);
        outputFile.close();
        workbook.close();
    }
    }
    
