package ru.nimdator;

import java.text.MessageFormat;
import java.util.LinkedList;
import java.util.List;
import java.util.regex.Pattern;

public class Main {
    public static void main(String[] args) {
        List<ExcelTestFile> fileContent = new LinkedList<>();
        try {
            ExcelMapper file = new ExcelMapper("test1.xlsx");
            fileContent = file.getFileContent(ExcelTestFile.class);
        } catch (ExcelFileMappingException e) {
            e.printStackTrace();
        }

        print("file {}", fileContent);
    }

    static void print(String s, Object... var2) {
        int i = 0;
        while(s.contains("{}")) {
            s = s.replaceFirst(Pattern.quote("{}"), "{"+ i++ +"}");
        }
        System.out.println(MessageFormat.format(s, var2));
    }
}