package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.font.FontRenderContext;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        String imagePath = "diploma_template.png"; // Numele imaginii diplomei goale
        String excelPath = "participanti.xlsx"; // Numele fișierului Excel
        String outputDir = "diplome_generate";

        new File(outputDir).mkdirs(); // Creează folderul de output

        try (FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0);
                if (cell != null) {
                    String name = cell.getStringCellValue();
                    int regNumber = row.getRowNum() + 1; // Număr de înregistrare
                    generateCertificate(imagePath, name, regNumber, outputDir);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void generateCertificate(String imagePath, String name, int regNumber, String outputDir) throws IOException {
        BufferedImage image = ImageIO.read(new File(imagePath));
        Graphics2D g = image.createGraphics();

        g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g.setColor(java.awt.Color.decode("#4257aa"));
        g.setFont(new java.awt.Font("Montserrat", java.awt.Font.BOLD, 90));

        // Calculează centrul imaginii
        int imageWidth = image.getWidth();
        FontMetrics fm = g.getFontMetrics();
        int textWidth = fm.stringWidth(name);
        int x = (imageWidth - textWidth) / 2 + 315; // Centrează textul
        int y = 1300; // Ajustează poziția pe verticală

        g.drawString(name, x, y);

        // Adaugă numărul de înregistrare
        g.setFont(new java.awt.Font("Montserrat", java.awt.Font.BOLD, 50));
        g.setColor(java.awt.Color.decode("#ffffff"));
        int x1 =2930; // Centrează textul
        int y1 = 2418; // Ajustează poziția pe verticală
        g.drawString(regNumber+"/190 ", x1, y1);

        g.dispose();

        File outputFile = new File(outputDir + "/" + name.replaceAll(" ", "_") + ".png");
        ImageIO.write(image, "png", outputFile);
    }
}
