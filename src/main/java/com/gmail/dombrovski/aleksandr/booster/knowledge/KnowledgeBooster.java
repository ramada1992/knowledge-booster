package com.gmail.dombrovski.aleksandr.booster.knowledge;

import java.util.*;
import java.util.stream.*;

public class KnowledgeBooster {
    public static void main(String[] argv) {
        System.out.println("Please enter file list in the following format FileName;FileName;FileName");
        Scanner input = new Scanner(System.in);
        String fileName = input.nextLine();
        String[] inputString;
        String separator = ";";
        inputString = fileName.split(separator);
        Stream<String> stream = Arrays.stream(inputString);
        stream.forEach(x -> System.out.println(GetMetaData.getMetaData(x)));

    }
}