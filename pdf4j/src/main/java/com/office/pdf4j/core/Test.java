package com.office.pdf4j.core;

import com.itextpdf.io.font.constants.StandardFonts;
import com.itextpdf.kernel.colors.ColorConstants;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.UnitValue;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;

public class Test {

    public static void main(String[] args) {
        File desc = new File("./target/sandbox/tables/simple_table.pdf");
        desc.getParentFile().mkdirs();
        try (PdfDocument pdfDoc = new PdfDocument(new PdfWriter(desc)); Document doc = new Document(pdfDoc)) {
            PdfFont pdfFont = PdfFontFactory.createFont(StandardFonts.HELVETICA_BOLD);
            Table table = new Table(UnitValue.createPercentArray(16)).useAllAvailableWidth();
            for (int i = 0; i < 16; i++) {
                Cell cell = new Cell().add(new Paragraph("hi"))
                        .setFont(pdfFont)
                        .setFontColor(ColorConstants.WHITE)
                        .setBackgroundColor(ColorConstants.BLUE)
                        .setBorder(Border.NO_BORDER)
                        .setTextAlignment(TextAlignment.CENTER);
                table.addCell(cell);
            }
            doc.add(table);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
