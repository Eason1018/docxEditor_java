package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;

import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class DocxUtils {

    /**
     * Analyzes the structure of a .docx file and prints table contents.
     * @param filePath path to the input .docx file
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
     * Adds a row to a given table.
     * @param table   the XWPFTable to add a row to
     * @param rowData the data for each cell in the new row
     * @return true if the row was added successfully
     */
    public static boolean addRowToTable(XWPFTable table, List<String> rowData) {
        XWPFTableRow newRow = table.createRow();
        List<XWPFTableCell> cells = newRow.getTableCells();

        for (int i = 0; i < rowData.size(); i++) {
            if (i < cells.size()) {
                cells.get(i).setText(rowData.get(i));
            } else {
                XWPFTableCell newCell = newRow.addNewTableCell();
                newCell.setText(rowData.get(i));
            }
        }
        return true;
    }

    /**
     * Deletes a row from a table by index.
     * @param table    the XWPFTable to delete a row from
     * @param rowIndex the index of the row to delete
     * @return true if the row was deleted successfully, false otherwise
     */
    public static boolean deleteRowFromTable(XWPFTable table, int rowIndex) {
        if (rowIndex < 0 || rowIndex >= table.getRows().size()) {
            return false;
        }
        table.removeRow(rowIndex);
        return true;
    }

    /**
     * Converts a .docx file to a PDF file.
     * Uses documents4j for conversion.
     *
     * @param inputDocxPath path to the input .docx file
     * @param outputPdfPath path to the output .pdf file
     */
    public static void convertToPdf(String inputDocxPath, String outputPdfPath) {
        File inputFile = new File(inputDocxPath);
        File outputFile = new File(outputPdfPath);

        try (FileInputStream inputStream = new FileInputStream(inputFile);
             FileOutputStream outputStream = new FileOutputStream(outputFile)) {

            // Create the converter
            IConverter converter = LocalConverter.builder().build();

            // Perform the conversion
            converter.convert(inputStream).as(DocumentType.MS_WORD)
                    .to(outputStream).as(DocumentType.PDF)
                    .execute();

            System.out.println("PDF created successfully at: " + outputPdfPath);

        } catch (Exception e) {
            System.err.println("Error during PDF conversion: " + e.getMessage());
        }
    }

}
