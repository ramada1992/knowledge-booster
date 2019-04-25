package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;


public class ExcelAuthorRemover {
    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String fileName) {
        final String inputFile = ("src/main/xls/" + fileName);

        try (final FileInputStream inputFileRead = new FileInputStream(inputFile)) {

            System.out.println("Remove authors name from " + fileName);

            final HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(inputFileRead));
            SummaryInformation summaryInfo = workbook.getSummaryInformation();
            String author = summaryInfo.getAuthor();

            if (author != null) {
                System.out.println("Current author name " + author);

                inputFileRead.close();
                summaryInfo.setAuthor(null);
                final FileOutputStream outputFileWrite = new FileOutputStream(inputFile);
                workbook.write(outputFileWrite);
                outputFileWrite.close();
            }
        } catch (IOException ex) {
            ex.getStackTrace();
        }
    }
}


