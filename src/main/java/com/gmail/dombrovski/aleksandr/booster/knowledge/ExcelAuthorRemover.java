package com.gmail.dombrovski.aleksandr.booster.knowledge;

import java.util.*;
import java.util.stream.*;

public class ExcelAuthorRemover {
    public static void main(String[] argv) {
            Stream<String> stream = Arrays.stream(argv);
            stream.forEach(x -> System.out.println(getMetaData(x)));
    }

    private static String getMetaData(String x){
        return x;
    }
}