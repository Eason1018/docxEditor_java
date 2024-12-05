package org.example;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.InputStream;

public class SignatureUtils {
    // Add a signature to a table cell
    public static void addSignatureToCell(XWPFTableCell cell, String imagePath) {
        try {
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            XWPFRun run = paragraph.createRun();

            try (InputStream is = new FileInputStream(imagePath)) {
                run.addPicture(is, Document.PICTURE_TYPE_PNG, imagePath, 100 * 9525, 50 * 9525);
                System.out.println("Signature added successfully.");
            }
        } catch (Exception e) {
            System.err.println("Error adding signature: " + e.getMessage());
        }
    }
}
