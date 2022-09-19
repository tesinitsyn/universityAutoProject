package org.example;

import java.io.IOException;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        ExcelParsing obj = new ExcelParsing();
        //get arrayList of Students and write file Path
        List<String> students = obj.pushToArrayList("C:\\Users\\sitie\\OneDrive\\Рабочий стол\\Пример таблицы.xlsx");
        //delete in future
        System.out.println(students);
    }
}