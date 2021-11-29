package com.example.demo.poi;

import java.io.FileInputStream;
import java.util.List;

public class poi {
    public poi() {

    }

    public void getExcelFileInfo() throws Exception {
        String filePath = "C:\\testFile.xlsx";

        FileInputStream fileInputStream = new FileInputStream(filePath);

        ExcelSheetHandler excelSheetHandler = ExcelSheetHandler.readExcel(fileInputStream);

        List<List<String>> excelDatas = excelSheetHandler.getRows();

        int iCol = 0;
        int iRow = 0;

        for(List<String> dataRow : excelDatas){
            for(String str : dataRow){
                if(iCol == 0){
                    //test@naver.com
                    System.out.println(str);
                }
                else if(iCol == 1){
                    //Seoul
                    System.out.println(str);
                }
                else if(iCol == 2){
                    //asv
                    System.out.println(str);
                }
                iCol++;
            }
            iCol = 0;
            iRow = 0;
        }
    }
}
