package com.gmail.dombrovski.aleksandr.booster.knowledge;

import java.util.*;
import java.util.stream.*;

public class ExcelAuthorRemover {
    public static void main(String[] argv) {
            Stream<String> stream = Arrays.stream(argv);
            stream.forEach(x -> removeAuthors(x));
    }

    private static void removeAuthors(String fileName){
        System.out.println("Remove authors name from "+ fileName );
    }
}