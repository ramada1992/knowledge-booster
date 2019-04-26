package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;



public class ExcelAuthorRemover {
    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String fileName) {
        System.out.println("Processing file " + fileName);
        final File file = new File(fileName);

        if (!file.exists()) {
            System.out.println("File doesn't exist");
            return;
        }

        if (fileName.toLowerCase().endsWith(".xls")) {
            processXLS(file);
        } else {
            System.err.println("Unsupported file type");
        }
    }

    private static void processXLS(final File file) {
        final HSSFWorkbook workbook = readXLS(file);
        final SummaryInformation summaryInfo = workbook.getSummaryInformation();
        final String author = summaryInfo.getAuthor();
        final String lastAuthor = summaryInfo.getLastAuthor();

        if (author != null && !author.isBlank()) {
            System.out.println("Cleaning author " + author);
            summaryInfo.setAuthor("");

            saveXLS(workbook, file);
        } else {
            System.out.println("Author already clean");
        }
        if (lastAuthor != null && !lastAuthor.isBlank()) {
            System.out.println("Cleaning last author " + lastAuthor);
            summaryInfo.setLastAuthor("");

            saveXLS(workbook, file);
        } else {
            System.out.println("Last author already clean");
        }
    }

    private static HSSFWorkbook readXLS(final File file) {
        try (FileInputStream input = new FileInputStream(file)) {
            return new HSSFWorkbook(new POIFSFileSystem(input));
        } catch (final IOException e) {
            throw new RuntimeException("Cannot read file", e);
        }
    }

    private static void saveXLS(final HSSFWorkbook workbook, final File file) {
        try {
            workbook.write(file);
        } catch (final IOException e) {
            throw new RuntimeException("Cannot save file", e);
        }
    }
}
