package com.gmail.dombrovski.aleksandr.booster.knowledge;

import java.util.*;
import java.util.stream.*;

public class KnowledgeBooster {
    public static void main(String[] argv) {

        if(argv == null){
            System.out.println("No files");
        } else {
            Stream<String> stream = Arrays.stream(argv);
            stream.forEach(x -> System.out.println(GetMetaData.getMetaData(x)));
        }

    }
}