package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;


public class ExcelAuthorRemover {
    public static void main(String[] fileName) {
        Arrays.stream(fileName).forEach(excelFile -> removeAuthors(excelFile));
    }

    private static void removeAuthors(String fileName) {
        System.out.println("Removing authors name from " + fileName);
        try {
            FileInputStream excelFile = new FileInputStream(new File("src/main/xls/" + fileName));
            Workbook workbook = new HSSFWorkbook(excelFile);
        } catch (IOException e) {
            System.err.format("IOException: %s%n", e);
        }
    }
}