package org.example;

import org.apache.poi.xwpf.usermodel.*;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;
import org.jodconverter.core.office.OfficeException;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import java.util.concurrent.TimeUnit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class DocxUtils {
    // Analyze the structure of a .docx file
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

    // Add a row to a table
    public static void addRowToTable(XWPFTable table, List<String> rowData) {
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
        System.out.println("Row added successfully.");
    }

    // Delete a row from a table
    public static void deleteRowFromTable(XWPFTable table, int rowIndex) {
        if (rowIndex < 0 || rowIndex >= table.getRows().size()) {
            System.out.println("Invalid row index.");
            return;
        }
        table.removeRow(rowIndex);
        System.out.println("Row deleted successfully.");
    }


    public static void convertToPdf(String inputDocxPath, String outputPdfPath) {
        File inputFile = new File(inputDocxPath);
        File outputFile = new File(outputPdfPath);

        try (FileInputStream inputStream = new FileInputStream(inputFile);
             FileOutputStream outputStream = new FileOutputStream(outputFile)) {

            IConverter converter = LocalConverter.builder()
                    .baseFolder(new File("temp")) // Temporary folder for conversion jobs
                    .workerPool(20, 25, 2, TimeUnit.SECONDS) // Configure worker threads
                    .build();

            boolean conversion = converter.convert(inputStream).as(DocumentType.MS_WORD)
                    .to(outputStream).as(DocumentType.PDF)
                    .execute();

            if (conversion) {
                System.out.println("PDF created successfully at: " + outputPdfPath);
            } else {
                System.err.println("PDF conversion failed.");
            }

        } catch (Exception e) {
            System.err.println("Error during PDF conversion: " + e.getMessage());
        }
    }
}