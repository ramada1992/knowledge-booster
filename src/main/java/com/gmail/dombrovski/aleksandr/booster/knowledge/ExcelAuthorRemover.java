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
    public static void main(String[] fileName) {
        Arrays.stream(fileName).forEach(excelFile -> removeAuthors(excelFile));
    }

    private static void removeAuthors(String fileName) {
        System.out.println("Removing authors name from " + fileName);
        try {

            FileInputStream stream = new FileInputStream("./src/main/xls/" + fileName);
            POIFSFileSystem poifs = new POIFSFileSystem(stream);
            DirectoryEntry dir = poifs.getRoot();
            DocumentEntry siEntry = (DocumentEntry) dir.getEntry(SummaryInformation.DEFAULT_STREAM_NAME);
            DocumentInputStream dis = new DocumentInputStream(siEntry);
            PropertySet ps = new PropertySet(dis);
            SummaryInformation si = new SummaryInformation(ps);
            String author = si.getAuthor();
            System.out.println(author);


        } catch (IOException | NoPropertySetStreamException | UnexpectedPropertySetTypeException ex) {
            ex.getStackTrace();
        }
    }
}