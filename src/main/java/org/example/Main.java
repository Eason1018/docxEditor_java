package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        String inputDocxPath = "src/main/resources/input.docx";
        try (Scanner scanner = new Scanner(System.in);
             XWPFDocument doc = new XWPFDocument(new FileInputStream(inputDocxPath))) {

            System.out.println("\nAnalyzing document contents...");
            DocxUtils.analyzeDocument(inputDocxPath);

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
            scanner.nextLine();
            return -1;
        }
    }

    private static void addRowInteraction(Scanner scanner, XWPFDocument doc) {
        try {
            System.out.print("Enter table index: ");
            int tableIndex = scanner.nextInt();
            if (!isValidTableIndex(tableIndex, doc)) return;

            System.out.print("Enter row data (comma-separated): ");
            scanner.nextLine();
            String rowData = scanner.nextLine();
            XWPFTable table = doc.getTables().get(tableIndex);
            DocxUtils.addRowToTable(table, Arrays.asList(rowData.split(",")));
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
            DocxUtils.deleteRowFromTable(table, rowIndex);
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
            scanner.nextLine();
            String[] indices = scanner.nextLine().split(",");
            if (indices.length < 2) {
                System.out.println("Invalid indices.");
                return;
            }
            int rowIndex = Integer.parseInt(indices[0].trim());
            int colIndex = Integer.parseInt(indices[1].trim());

            System.out.print("Enter image path: ");
            String imagePath = scanner.nextLine();

            XWPFTable table = doc.getTables().get(tableIndex);
            if (rowIndex < 0 || rowIndex >= table.getNumberOfRows()) return;
            XWPFTableRow row = table.getRow(rowIndex);
            if (colIndex < 0 || colIndex >= row.getTableCells().size()) return;

            XWPFTableCell cell = row.getCell(colIndex);
            SignatureUtils.addSignatureToCell(cell, imagePath);
        } catch (Exception e) {
            System.out.println("Error adding signature: " + e.getMessage());
        }
    }

    private static boolean isValidTableIndex(int tableIndex, XWPFDocument doc) {
        return tableIndex >= 0 && tableIndex < doc.getTables().size();
    }

    private static void saveAndConvert(XWPFDocument doc) {
        String docxPath = "./resources/output.docx";
        String pdfPath = "./resources/output.pdf";

        try (FileOutputStream out = new FileOutputStream(docxPath)) {
            doc.write(out);
            System.out.println("Document saved as " + docxPath);
            DocxUtils.convertToPdf(docxPath, pdfPath);
            System.out.println("PDF saved as " + pdfPath);
        } catch (Exception e) {
            System.err.println("Error saving document: " + e.getMessage());
        }
    }
}
