package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;

public class ExcelAuthorRemover {
    private static final Map<Class<? extends Workbook>, Function<Workbook, String>> GET_AUTHOR = Map.of(
            HSSFWorkbook.class, (document) ->
                    documentSummary(document).getAuthor(),
            XSSFWorkbook.class, (document) ->
                    documentProperties(document).getCreator());
    private static final Map<Class<? extends Workbook>, BiConsumer<Workbook, String>> SET_AUTHOR = Map.of(
            HSSFWorkbook.class, (document, author) ->
                    documentSummary(document).setAuthor(author),
            XSSFWorkbook.class, (document, author) ->
                    documentProperties(document).setCreator(author));
    private static final Map<Class<? extends Workbook>, Function<Workbook, String>> GET_LAST_AUTHOR = Map.of(
            HSSFWorkbook.class, (document) ->
                    documentSummary(document).getLastAuthor(),
            XSSFWorkbook.class, (document) ->
                    documentProperties(document).getLastModifiedByUser());
    private static final Map<Class<? extends Workbook>, BiConsumer<Workbook, String>> SET_LAST_AUTHOR = Map.of(
            HSSFWorkbook.class, (document, lastAuthor) ->
                    documentSummary(document).setLastAuthor(lastAuthor),
            XSSFWorkbook.class, (document, lastAuthor) ->
                    documentProperties(document).setLastModifiedByUser(lastAuthor));

    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String filename) {
        System.out.println("Cleaning file " + filename);

        try {
            final Workbook document = readDocument(filename);
            final String author = getDocumentProperty(document, GET_AUTHOR);
            final String lastAuthor = getDocumentProperty(document, GET_LAST_AUTHOR);

            boolean removed = false;

            if (author != null && !author.isBlank()) {
                System.out.println("Cleaning author " + author);
                cleanDocumentProperty(document, SET_AUTHOR);
                removed = true;
            }

            if (lastAuthor != null && !lastAuthor.isBlank()) {
                System.out.println("Cleaning last author " + lastAuthor);
                cleanDocumentProperty(document, SET_LAST_AUTHOR);
                removed = true;
            }

            if (removed) {
                saveDocument(document, filename);
            } else {
                System.out.println("Already clean");
            }
        } catch (final Exception e) {
            System.out.println(fullErrorMessage(e));
        }
    }

    private static Workbook readDocument(final String filename) {
        try (FileInputStream stream = new FileInputStream(filename)) {
            return WorkbookFactory.create(stream);
        } catch (final IOException e) {
            throw new RuntimeException("Cannot read Excel document", e);
        }
    }

    private static void saveDocument(final Workbook workbook, final String filename) {
        try (FileOutputStream stream = new FileOutputStream(filename)) {
            workbook.write(stream);
        } catch (final IOException e) {
            throw new RuntimeException("Cannot save Excel document", e);
        }
    }

    private static String getDocumentProperty(final Workbook document,
                                              final Map<Class<? extends Workbook>,
                                                      Function<Workbook, String>> getters) {
        final Function<Workbook, String> getter = getters.get(document.getClass());
        if (getter == null) {
            throw new RuntimeException("Unsupported Excel document type: " + document.getClass());
        }

        return getter.apply(document);
    }

    private static void cleanDocumentProperty(final Workbook document,
                                              final Map<Class<? extends Workbook>,
                                                      BiConsumer<Workbook, String>> setters) {
        final BiConsumer<Workbook, String> setter = setters.get(document.getClass());
        if (setter == null) {
            throw new RuntimeException("Unsupported Excel document type: " + document.getClass());
        }

        setter.accept(document, "");
    }

    private static SummaryInformation documentSummary(final Workbook document) {
        return ((HSSFWorkbook) document).getSummaryInformation();
    }

    private static CoreProperties documentProperties(final Workbook document) {
        return ((XSSFWorkbook) document).getProperties().getCoreProperties();
    }

    private static String fullErrorMessage(final Throwable error) {
        final StringBuilder message = new StringBuilder();
        for (Throwable current = error; current != null; current = current.getCause()) {
            if (message.length() > 0) {
                message.append(": ");
            }
            message.append(current.getMessage());

            if (current == current.getCause()) {
                break;
            }
        }

        return message.toString();
    }
}
