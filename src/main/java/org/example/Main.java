package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.Scanner;

public class Main {
    public static void main(String[] args) {
        try (XWPFDocument doc = new XWPFDocument(new FileInputStream("src/main/resources/input.docx"))) { // Load the input .docx file
            Scanner scanner = new Scanner(System.in); // Initialize a Scanner for user input

            System.out.println("Choose an action: ");
            System.out.println("1. Analyze Document");
            System.out.println("2. Add a Row");
            System.out.println("3. Add a Signature");
            int choice = scanner.nextInt(); // Get the user's choice

            switch (choice) {
                case 1:
                    // Analyze the document structure
                    DocxUtils.analyzeDocument("input.docx");
                    break;

                case 2:
                    // Add a row to a specified table
                    System.out.print("Enter table index: ");
                    int tableIndex = scanner.nextInt(); // Get the index of the table
                    System.out.print("Enter row data (comma-separated): ");
                    scanner.nextLine(); // Consume newline character
                    String rowData = scanner.nextLine(); // Get the row data from the user
                    XWPFTable table = doc.getTables().get(tableIndex); // Access the specified table
                    DocxUtils.addRowToTable(table, Arrays.asList(rowData.split(","))); // Add the row
                    break;

                case 3:
                    // Add a signature image to a table cell
                    System.out.print("Enter table index: ");
                    int sigTableIndex = scanner.nextInt(); // Get the index of the table
                    System.out.print("Enter row and column index (comma-separated): ");
                    scanner.nextLine(); // Consume newline character
                    String[] indices = scanner.nextLine().split(","); // Get row and column indices
                    int rowIndex = Integer.parseInt(indices[0]); // Parse row index
                    int colIndex = Integer.parseInt(indices[1]); // Parse column index
                    System.out.print("Enter image path: ");
                    String imagePath = scanner.nextLine(); // Get the path to the image

                    // Access the specified cell
                    XWPFTableCell cell = doc.getTables().get(sigTableIndex).getRow(rowIndex).getCell(colIndex);
                    SignatureUtils.addSignatureToCell(cell, imagePath); // Add the signature
                    break;

                default:
                    System.out.println("Invalid choice."); // Handle invalid user input
            }

            // Save changes to the output .docx file
            try (FileOutputStream out = new FileOutputStream("output.docx")) {
                doc.write(out); // Write the document to the output file
            }
            System.out.println("Document saved as output.docx"); // Log the save operation
        } catch (Exception e) {
            System.err.println("Error: " + e.getMessage()); // Handle errors gracefully
        }
    }
}
