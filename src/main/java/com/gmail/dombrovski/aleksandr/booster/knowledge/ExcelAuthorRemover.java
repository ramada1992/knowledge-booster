package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Function;
import java.util.stream.Stream;

public class ExcelAuthorRemover {
    private static final Logger LOGGER = LoggerFactory.getLogger(ExcelAuthorRemover.class);

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
        LOGGER.info("Cleaning file: {}", filename);

        try {
            final Workbook document = readDocument(filename);
            if (Stream.of(
                    Pair.create("author", Pair.create(GET_AUTHOR, SET_AUTHOR)),
                    Pair.create("lastAuthor", Pair.create(GET_LAST_AUTHOR, SET_LAST_AUTHOR)))
                    .map(property -> cleanProperty(document, property.getKey(),
                            property.getValue().getFirst(), property.getValue().getSecond()))
                    .reduce(false, (a, b) -> a || b)) {
                saveDocument(document, filename);
            } else {
                LOGGER.info("Already clean");
            }
        } catch (final Exception e) {
            LOGGER.error("Cannot clean file: {}", filename, e);
        }
    }

    private static boolean cleanProperty(final Workbook document,
                                         final String name,
                                         final Map<Class<? extends Workbook>, Function<Workbook, String>> getters,
                                         final Map<Class<? extends Workbook>, BiConsumer<Workbook, String>> setters) {
        final String property = getDocumentProperty(document, getters);
        if (StringUtils.isNotBlank(property)) {
            LOGGER.info("Cleaning {}: {}", name, property);
            cleanDocumentProperty(document, setters);
            return true;
        }

        return false;
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
}

