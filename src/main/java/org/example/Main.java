package org.example;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        ExcelParsing objForExcelParsing = new ExcelParsing();
        //get arrayList of Students and write file Path
        List<String> students = objForExcelParsing.pushToArrayList("C:\\Users\\sitie\\OneDrive\\Рабочий стол\\Пример таблицы.xlsx");
        //delete in future
        System.out.println(students);
        ArrayList<String> replaceableNames = new ArrayList<>();
        replaceableNames.add("${instituteName}");
        replaceableNames.add("${departmentName}");
        replaceableNames.add("${practiceName}");
        replaceableNames.add("${orderDate}");
        replaceableNames.add("${orderName}");
        replaceableNames.add("${sessionDate}");
        replaceableNames.add("${studentFN}");
        replaceableNames.add("${supervisorFN}");
        replaceableNames.add("${currentYear}");
        replaceableNames.add("${courseNum}");
        replaceableNames.add("${groupName}");
        replaceableNames.add("${studentFullName}");
        replaceableNames.add("${practicePlaceAndTime}");
        replaceableNames.add("${position}");
        replaceableNames.add("${currentDate}");
        replaceableNames.add("${headOfDFN}");
        replaceableNames.add("${directionNum}");
        replaceableNames.add("${directionName}");
        replaceableNames.add("${profileName}");
        UpdateDocument objUpdateWord = new UpdateDocument();
        String inputPath = "D:\\wrdfiles\\Титулы практики.docx";
        String outputPath = "D:\\wrdfiles\\projecttest.docx";
        objUpdateWord.updateDocument(inputPath, outputPath, "Морозов", "Солянов");
    }
}