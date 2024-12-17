package org.example;

import org.jodconverter.core.DocumentConverter;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

public class DocxUtils {

    /**
     * Analyzes the structure of a .docx file and prints table contents.
     * @param filePath Path to the input .docx file
     * @throws IOException if an I/O error occurs
     */
    public static void analyzeDocument(String filePath) throws IOException {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream(filePath))) {
            System.out.println("Analyzing document: " + filePath);

            int tableIndex = 0;
            for (XWPFTable table : doc.getTables()) {
                System.out.println("\nTable " + tableIndex + ":");
                int rowIndex = 0;
                for (XWPFTableRow row : table.getRows()) {
                    System.out.print("Row " + rowIndex + ": ");
                    for (XWPFTableCell cell : row.getTableCells()) {
                        System.out.print(cell.getText().strip() + " | ");
                    }
                    System.out.println();
                    rowIndex++;
                }
                tableIndex++;
            }
        }
    }

    /**
     * Converts a DOCX file to a PDF file using LibreOffice.
     * @param inputDocxPath Path to the input DOCX file
     * @param outputPdfPath Path to the output PDF file
     */
    public static void convertToPdf(String inputDocxPath, String outputPdfPath) {
        LocalOfficeManager officeManager = LocalOfficeManager.builder().install().build();
        try {
            // Start LibreOffice service
            officeManager.start();

            // Create the converter
            DocumentConverter converter = LocalConverter.make(officeManager);

            // Perform the conversion
            converter.convert(new File(inputDocxPath))
                    .to(new File(outputPdfPath))
                    .execute();

            System.out.println("PDF created successfully at: " + outputPdfPath);

        } catch (Exception e) {
            System.err.println("Error during PDF conversion: " + e.getMessage());
        } finally {
            try {
                officeManager.stop();
            } catch (Exception e) {
                System.err.println("Error stopping LibreOffice manager: " + e.getMessage());
            }
        }
    }

    /**
     * Adds a new row to the specified table.
     * @param table   The table to add a row to
     * @param rowData The data for the new row
     * @return True if successful
     */
    public static boolean addRowToTable(XWPFTable table, List<String> rowData) {
        XWPFTableRow newRow = table.createRow();
        List<XWPFTableCell> cells = newRow.getTableCells();

        // Ensure the row has enough cells
        while (cells.size() < rowData.size()) {
            newRow.addNewTableCell();
        }

        // Populate cells
        for (int i = 0; i < rowData.size(); i++) {
            cells.get(i).setText(rowData.get(i));
        }
        return true;
    }

    /**
     * Deletes a row from the specified table.
     * @param table    The table to delete a row from
     * @param rowIndex The index of the row to delete
     * @return True if successful
     */
    public static boolean deleteRowFromTable(XWPFTable table, int rowIndex) {
        if (rowIndex < 0 || rowIndex >= table.getRows().size()) return false;
        table.removeRow(rowIndex);
        return true;
    }
}
