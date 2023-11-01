package com.codeforyou.apachepoi;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
public class ExcelController {

    @GetMapping("/excel")
    public void excelGenerator() throws Exception {

        Map<String, String> jsonToExcelMapping = new HashMap<>();
        jsonToExcelMapping.put("tripuratenders.gov.in", "Tripura");
        jsonToExcelMapping.put("mptenders.gov.in", "Madhya Pradesh");
        jsonToExcelMapping.put("eprocure.goa.gov.in", "Goa");
        jsonToExcelMapping.put("etender.up.nic.in", "Uttar Pradesh");
        jsonToExcelMapping.put("coalindiatenders.nic.in", "Coal India Limited");
        jsonToExcelMapping.put("manipurtenders.gov.in", "Manipur");
        jsonToExcelMapping.put("uktenders.gov.in", "Uttarakhand");
        jsonToExcelMapping.put("eprocure.andaman.gov.in", "Andaman & Nicobar");
        jsonToExcelMapping.put("eprocurentpc.nic.in", "NTPC");
        jsonToExcelMapping.put("jharkhandtenders.gov.in", "Jharkhand");
        jsonToExcelMapping.put("mahatenders.gov.in", "Maharashtra");
        jsonToExcelMapping.put("dnhtenders.gov.in", "DadraNH");
        jsonToExcelMapping.put("eproc.rajasthan.gov.in", "Rajasthan");
        jsonToExcelMapping.put("eprocurebhel.co.in", "Barath Heavy Electricals Limited");
        jsonToExcelMapping.put("eproc.punjab.gov.in", "Punjab PWD");
        jsonToExcelMapping.put("etenders.kerala.gov.in", "Kerala");
        jsonToExcelMapping.put("assamtenders.gov.in", "Assam");
        jsonToExcelMapping.put("etenders.hry.nic.in", "Haryana");
        jsonToExcelMapping.put("meghalayatenders.gov.in", "Meghalaya");
        jsonToExcelMapping.put("iocletenders.nic.in", "Indian Oil Corp Ltd");
        jsonToExcelMapping.put("tntenders.gov.in", "Tamil Nadu");
        jsonToExcelMapping.put("pudutenders.gov.in", "Puducherry");
        jsonToExcelMapping.put("arunachaltenders.gov.in", "Arunachal Pradesh");
        jsonToExcelMapping.put("hptenders.gov.in", "Himachal Pradesh");
        jsonToExcelMapping.put("mizoramtenders.gov.in", "Mizoram");
        jsonToExcelMapping.put("ddtenders.gov.in", "DamnDiu");
        jsonToExcelMapping.put("nagalandtenders.gov.in", "Nagaland");
        jsonToExcelMapping.put("sikkimtender.gov.in", "Sikkim");
        jsonToExcelMapping.put("jktenders.gov.in", "Jammu and Kashmir");
        jsonToExcelMapping.put("govtprocurement.delhi.gov.in", "Delhi NCT");
        jsonToExcelMapping.put("tendersutl.gov.in", "Lakshadweep");
        jsonToExcelMapping.put("wbtenders.gov.in", "West Bengal");
        jsonToExcelMapping.put("tendersodisha.gov.in", "Odissa");
        jsonToExcelMapping.put("etenders.chd.nic.in", "Chandigarh");
        
        FileInputStream templateFile = new FileInputStream("template.xls");
        Workbook workbook = WorkbookFactory.create(templateFile);
        templateFile.close();

        Sheet sheet = workbook.getSheetAt(0);

        String searchString = "Central CPSEs";

        String valueToInsert = "678";

        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING) {
                    String cellValue = cell.getStringCellValue();
                    if (cellValue.contains(searchString)) {
                        System.out.println("Found '" + searchString + "' in cell: " + cell.getAddress());
                        
                        int nextColumnIndex = cell.getColumnIndex() + 1;
                        Cell nextCell = row.createCell(nextColumnIndex);
                        nextCell.setCellValue(valueToInsert);
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
