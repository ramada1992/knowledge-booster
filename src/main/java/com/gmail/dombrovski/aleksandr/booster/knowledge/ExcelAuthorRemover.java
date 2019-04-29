package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

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
        final File file = new File(fileName);

        System.out.println("Processing file " + file.getName());

        if (!file.exists()) {
            System.out.println("File doesn't exist");
            return;
        }
        if (fileName.toLowerCase().endsWith(".xls")) {
            processXLS(file);
        }
        if (fileName.toLowerCase().endsWith(".xlsx")) {
            processXLSX(file);
        } else {
            System.out.println("Unsupported file type");
        }
    }

    //***       Start_Logic - Process XLS documents Microsoft Office 1997-2004        ***

    private static void processXLS(final File file) {
        final HSSFWorkbook workbook = readXLS(file);
        final SummaryInformation summaryInfo = workbook.getSummaryInformation();
        final String author = summaryInfo.getAuthor();
        final String lastAuthor = summaryInfo.getLastAuthor();

        boolean flagSaveFile = false;

        if (author != null && !author.isBlank()) {
            System.out.println("Cleaning author " + author);
            summaryInfo.setAuthor("");
            flagSaveFile = true;
        } else {
            System.out.println("Author already clean");
        }
        if (lastAuthor != null && !lastAuthor.isBlank()) {
            System.out.println("Cleaning last author " + lastAuthor);
            summaryInfo.setLastAuthor("");
            flagSaveFile = true;
        } else {
            System.out.println("Last author already clean");
        }
        if (flagSaveFile) {
            saveXLS(workbook, file);
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

    //***       Start_Logic - Process XLSX documents Microsoft Office 2005 to current        ***

    private static void processXLSX(final File file) {
        final XSSFWorkbook workbook = readXLSX(file);
        final POIXMLProperties props = workbook.getProperties();
        final POIXMLProperties.CoreProperties coreProp = props.getCoreProperties();
        final String author = coreProp.getCreator();
        final String lastAuthor = coreProp.getLastModifiedByUser();

        boolean flagSaveFile = false;

        if (author != null && !author.isBlank()) {
            System.out.println("Cleaning author " + author);
            coreProp.setCreator("");
            flagSaveFile = true;
        } else {
            System.out.println("Author already clean");
        }
        if (lastAuthor != null && !lastAuthor.isBlank()) {
            System.out.println("Cleaning last author " + lastAuthor);
            coreProp.setLastModifiedByUser("");
            flagSaveFile = true;
        } else {
            System.out.println("Last author already clean");
        }
        if (flagSaveFile) {
            saveXLSX(workbook, file);
        }
    }

    private static XSSFWorkbook readXLSX(final File file) {
        try (FileInputStream input = new FileInputStream(file)) {
            return new XSSFWorkbook(input);
        } catch (final IOException e) {
            throw new RuntimeException("Cannot read file", e);
        }
    }

    private static void saveXLSX(final XSSFWorkbook workbook, final File file) {
        try (FileOutputStream output = new FileOutputStream(file)) {
            workbook.write(output);
        } catch (final IOException e) {
            throw new RuntimeException("Cannot save file", e);
        }
    }

}
