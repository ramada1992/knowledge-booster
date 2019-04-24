package com.gmail.dombrovski.aleksandr.booster.knowledge;

import org.apache.poi.hpsf.NoPropertySetStreamException;
import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.hpsf.UnexpectedPropertySetTypeException;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Arrays;


public class ExcelAuthorRemover {
    public static void main(final String[] fileName) {
        Arrays.stream(fileName).forEach(ExcelAuthorRemover::removeAuthors);
    }

    private static void removeAuthors(final String fileName) {
        try (FileInputStream stream = new FileInputStream("src/main/xls/" + fileName)) {

            System.out.println("Remove authors name from " + fileName);

            DirectoryEntry dir = new POIFSFileSystem(stream).getRoot();
            DocumentEntry siEntry = (DocumentEntry) dir.getEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            PropertySet ps = new PropertySet(new DocumentInputStream(siEntry));
            String author = new SummaryInformation(ps).getAuthor();

            System.out.println("File author " + author);

        } catch (IOException | NoPropertySetStreamException | UnexpectedPropertySetTypeException ex) {
            ex.getStackTrace();
        }
    }
}