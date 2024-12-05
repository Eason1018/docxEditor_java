package org.example;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.util.List;

public class DocxUtils {
    // Analyze the structure of a .docx file by printing out its table contents
    public static void analyzeDocument(String filePath) throws IOException {
        XWPFDocument doc = new XWPFDocument(new FileInputStream(filePath)); // Open the .docx file as an XWPFDocument
        System.out.println("Analyzing document: " + filePath);

        int tableIndex = 0;
        for (XWPFTable table : doc.getTables()) { // Loop through all tables in the document
            System.out.println("\nTable " + tableIndex + ":");
            int rowIndex = 0;
            for (XWPFTableRow row : table.getRows()) { // Loop through each row in the table
                System.out.print("Row " + rowIndex + ": ");
                for (XWPFTableCell cell : row.getTableCells()) { // Loop through each cell in the row
                    System.out.print(cell.getText().strip() + " | "); // Print the text content of the cell
                }
                System.out.println(); // Move to the next row
                rowIndex++;
            }
            tableIndex++;
        }
    }

    // Add a new row to a table and populate it with data
    public static void addRowToTable(XWPFTable table, List<String> rowData) {
        XWPFTableRow newRow = table.createRow(); // Create a new row at the end of the table
        List<XWPFTableCell> cells = newRow.getTableCells(); // Get all cells in the new row

        for (int i = 0; i < rowData.size(); i++) {
            if (i < cells.size()) { // If the table already has enough cells
                cells.get(i).setText(rowData.get(i)); // Set text for existing cells
            } else {
                XWPFTableCell newCell = newRow.addNewTableCell(); // Create a new cell if needed
                newCell.setText(rowData.get(i)); // Populate it with the corresponding data
            }
        }
        System.out.println("Row added successfully."); // Log the success of the operation
    }
}
