package com.gmail.dombrovski.aleksandr.booster.knowledge;

import java.util.*;
import java.util.stream.*;

public class KnowledgeBooster {
    public static void main(String[] argv) {

        if(argv.length == 0){
            System.out.println("No files");
        } else {
            String separator = ";";
            String inputString = argv[0];
            String[] inputString2 = inputString.split(separator);
            Stream<String> stream = Arrays.stream(inputString2);
            stream.forEach(x -> System.out.println(GetMetaData.getMetaData(x)));
        }

    }
}