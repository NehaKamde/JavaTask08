package javapackage;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class FilelWriteOperation {
    public static void main(String[] args) {
        // Data to write
        String[][] data = {
            {"Name", "Age", "Email"},
            {"John Doe", "30", "john@test.com"},
            {"Jane Doe", "28", "john@test.com"},
            {"Bob Smith", "35", "jacky@example.com"},
            {"Swapnil", "37", "swapnil@example.com"}
        };

        // Create a new workbook and sheet
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");

        // Populate sheet with data
        for (int i = 0; i < data.length; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < data[i].length; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(data[i][j]);
            }
        }

        // Write to Excel file
        try (FileOutputStream fos = new FileOutputStream("data.xlsx")) {
            workbook.write(fos);
            System.out.println("Excel file 'data.xlsx' written successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }

        // Close the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
