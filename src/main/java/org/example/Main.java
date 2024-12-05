package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream("src/main/resources/input.docx"))) {
            Scanner scanner = new Scanner(System.in);

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
                int choice = scanner.nextInt();

                switch (choice) {
                    case 1 -> {
                        System.out.print("Enter table index: ");
                        int tableIndex = scanner.nextInt();
                        System.out.print("Enter row data (comma-separated): ");
                        scanner.nextLine(); // Consume newline
                        String rowData = scanner.nextLine();
                        XWPFTable table = doc.getTables().get(tableIndex);
                        DocxUtils.addRowToTable(table, Arrays.asList(rowData.split(",")));
                    }
                    case 2 -> {
                        System.out.print("Enter table index: ");
                        int tableIndex = scanner.nextInt();
                        System.out.print("Enter row index to delete: ");
                        int rowIndex = scanner.nextInt();
                        XWPFTable table = doc.getTables().get(tableIndex);
                        DocxUtils.deleteRowFromTable(table, rowIndex);
                    }
                    case 3 -> {
                        System.out.print("Enter table index: ");
                        int tableIndex = scanner.nextInt();
                        System.out.print("Enter row and column index (comma-separated): ");
                        scanner.nextLine(); // Consume newline
                        String[] indices = scanner.nextLine().split(",");
                        int rowIndex = Integer.parseInt(indices[0]);
                        int colIndex = Integer.parseInt(indices[1]);
                        System.out.print("Enter image path: ");
                        String imagePath = scanner.nextLine();
                        XWPFTableCell cell = doc.getTables().get(tableIndex).getRow(rowIndex).getCell(colIndex);
                        SignatureUtils.addSignatureToCell(cell, imagePath);
                    }
                    case 4 -> {
                        // Save document and output as PDF
                        running = false;
                        System.out.println("Saving document and generating PDF...");
                        try (FileOutputStream out = new FileOutputStream("src/main/resources/output.docx")) {
                            doc.write(out);
                        }
                        System.out.println("Document saved as output.docx.");

                        // Convert to PDF using LibreOffice
                        try {
                            DocxUtils.convertToPdf("src/main/resources/output.docx", "src/main/resources/output.pdf");
                            System.out.println("PDF saved as output.pdf.");
                        } catch (Exception e) {
                            System.err.println("Error during PDF conversion: " + e.getMessage());
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage());
        }
    }
}
