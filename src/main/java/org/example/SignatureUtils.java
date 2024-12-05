package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;

public class SignatureUtils {
    // Add a signature image to a specific table cell
    public static void addSignatureToCell(XWPFTableCell cell, String imagePath) {
        try {
            XWPFParagraph paragraph = cell.getParagraphs().get(0); // Get the first paragraph in the cell
            XWPFRun run = paragraph.createRun(); // Create a run to hold the image

            try (InputStream is = new FileInputStream(imagePath)) { // Open the image file as an InputStream
                // Add the image to the run, specifying its type and dimensions
                run.addPicture(is, Document.PICTURE_TYPE_PNG, imagePath, 100 * 9525, 50 * 9525); // Width: 100pt, Height: 50pt
                System.out.println("Signature added successfully."); // Log the success of the operation
            }
        } catch (Exception e) {
            System.err.println("Error adding signature: " + e.getMessage()); // Print any errors encountered
        }
    }
}
