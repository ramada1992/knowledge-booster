package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;


public class ExcelAuthorRemover {
    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String fileName) {
        System.out.println("Remove author from file" + fileName);

        File file = new File(fileName);

        if (FilenameUtils.getExtension(fileName).equals("xls")) {
            System.out.println("Remove authors name from " + fileName);

            processXLS(fileName);
        }
    }

    private static void processXLS(String fileName) {

        try (final FileInputStream inputFileRead = new FileInputStream(fileName)) {
            final HSSFWorkbook workbook = new HSSFWorkbook(new POIFSFileSystem(inputFileRead));
            final SummaryInformation summaryInfo = workbook.getSummaryInformation();
            final String author = summaryInfo.getAuthor();

            if (author != null) {
                System.out.println("Current author name " + author);

                inputFileRead.close();
                //summaryInfo.setAuthor("New");
                final FileOutputStream outputFileWrite = new FileOutputStream(fileName);
                workbook.write(outputFileWrite);
                outputFileWrite.close();
            }

        } catch (IOException ex) {
            ex.getStackTrace();
        }

    }

}




