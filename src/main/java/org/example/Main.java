package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        try (Scanner scanner = new Scanner(System.in);
             XWPFDocument doc = new XWPFDocument(new FileInputStream("src/main/resources/input.docx"))) {

            System.out.println("\nAnalyzing document contents...");
            DocxUtils.analyzeDocument("src/main/resources/input.docx");

            boolean running = true;
            while (running) {
                System.out.println("\nModification Menu:");
                System.out.println("1. Add a Row");
                System.out.println("2. Delete a Row");
                System.out.println("3. Add a Signature");
                System.out.println("4. No further modifications (Quit and output to PDF)");
                System.out.print("Choose an action (1-4): ");

                int choice = getUserChoice(scanner);
                switch (choice) {
                    case 1 -> addRowInteraction(scanner, doc);
                    case 2 -> deleteRowInteraction(scanner, doc);
                    case 3 -> addSignatureInteraction(scanner, doc);
                    case 4 -> {
                        running = false;
                        saveAndConvert(doc);
                    }
                    default -> System.out.println("Invalid choice. Please choose between 1-4.");
                }
            }
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
        }
    }

    private static int getUserChoice(Scanner scanner) {
        try {
            return scanner.nextInt();
        } catch (Exception e) {
            System.out.println("Invalid input. Please enter a number (1-4).");
            scanner.nextLine(); // consume invalid input
            return -1;
        }
    }

    private static void addRowInteraction(Scanner scanner, XWPFDocument doc) {
        try {
            System.out.print("Enter table index: ");
            int tableIndex = scanner.nextInt();
            if (!isValidTableIndex(tableIndex, doc)) return;

            System.out.print("Enter row data (comma-separated): ");
            scanner.nextLine(); // consume newline
            String rowData = scanner.nextLine();
            XWPFTable table = doc.getTables().get(tableIndex);
            DocxUtils.addRowToTable(table, Arrays.asList(rowData.split(","))); // Void method; no need for boolean
            System.out.println("Row added successfully.");
        } catch (Exception e) {
            System.out.println("Error adding row: " + e.getMessage());
        }
    }

    private static void deleteRowInteraction(Scanner scanner, XWPFDocument doc) {
        try {
            System.out.print("Enter table index: ");
            int tableIndex = scanner.nextInt();
            if (!isValidTableIndex(tableIndex, doc)) return;

            System.out.print("Enter row index to delete: ");
            int rowIndex = scanner.nextInt();
            XWPFTable table = doc.getTables().get(tableIndex);
            DocxUtils.deleteRowFromTable(table, rowIndex); // Void method; no need for boolean
            System.out.println("Row deleted successfully.");
        } catch (Exception e) {
            System.out.println("Error deleting row: " + e.getMessage());
        }
    }

    private static void addSignatureInteraction(Scanner scanner, XWPFDocument doc) {
        try {
            System.out.print("Enter table index: ");
            int tableIndex = scanner.nextInt();
            if (!isValidTableIndex(tableIndex, doc)) return;

            System.out.print("Enter row and column index (comma-separated): ");
            scanner.nextLine(); // consume newline
            String[] indices = scanner.nextLine().split(",");
            if (indices.length < 2) {
                System.out.println("Invalid indices. Please provide row and column separated by a comma.");
                return;
            }
            int rowIndex = Integer.parseInt(indices[0].trim());
            int colIndex = Integer.parseInt(indices[1].trim());

            System.out.print("Enter image path: ");
            String imagePath = scanner.nextLine();

            XWPFTable table = doc.getTables().get(tableIndex);
            if (rowIndex < 0 || rowIndex >= table.getNumberOfRows()) {
                System.out.println("Invalid row index.");
                return;
            }
            XWPFTableRow row = table.getRow(rowIndex);
            if (colIndex < 0 || colIndex >= row.getTableCells().size()) {
                System.out.println("Invalid column index.");
                return;
            }

            XWPFTableCell cell = row.getCell(colIndex);
            SignatureUtils.addSignatureToCell(cell, imagePath); // Void method; no need for boolean
            System.out.println("Signature added successfully.");
        } catch (Exception e) {
            System.out.println("Error adding signature: " + e.getMessage());
        }
    }

    private static boolean isValidTableIndex(int tableIndex, XWPFDocument doc) {
        if (tableIndex < 0 || tableIndex >= doc.getTables().size()) {
            System.out.println("Invalid table index.");
            return false;
        }
        return true;
    }

    private static void saveAndConvert(XWPFDocument doc) {
        String docxPath = "C:\\Temp\\output.docx";
        String pdfPath = "C:\\Temp\\output.pdf";

        System.out.println("Saving document and generating PDF...");
        try (FileOutputStream out = new FileOutputStream(docxPath)) {
            doc.write(out);
            System.out.println("Document saved as " + docxPath);
        } catch (Exception e) {
            System.err.println("Error saving document: " + e.getMessage());
        }

        try {
            DocxUtils.convertToPdf(docxPath, pdfPath);
            System.out.println("Document converted to PDF successfully.");
        } catch (Exception e) {
            System.err.println("Error during PDF conversion: " + e.getMessage());
        }
    }
}
