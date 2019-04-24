package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;


public class ExcelAuthorRemover {
    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String fileName) {
        try (FileInputStream fileStream = new FileInputStream("src/main/xls/" + fileName)) {

            System.out.println("Remove authors name from " + fileName);

            final HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(fileStream));
            SummaryInformation summaryInfo = workbook.getSummaryInformation();

            System.out.println("File author " + summaryInfo.getAuthor());

        } catch (IOException ex) {
            ex.getStackTrace();
        }
    }
}