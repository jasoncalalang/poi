package application;

import javax.ws.rs.GET;
import javax.ws.rs.POST;
import javax.ws.rs.Path;
import javax.ws.rs.PathParam;
import javax.ws.rs.core.Context;
import javax.ws.rs.Produces;
import javax.ws.rs.core.Application;
import javax.ws.rs.core.Context;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import javax.ws.rs.core.Request;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import java.io.File;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import javax.servlet.http.HttpServletRequest;
import java.util.Enumeration;
import java.io.FileOutputStream;
import org.apache.poi.xslf.usermodel.*;
import java.awt.Color;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import java.awt.Rectangle;

@Path("upload")
public class UploadEndpoint {
    public static final String SAMPLE_XLSX_FILE_PATH = "/home/relucio/microservicebuilder/poi/spreadsheet.xlsx";

    @POST
    @Produces(MediaType.APPLICATION_JSON)
    public Response uploadxls(@Context Request request) {
        System.out.println("we are here");
        try {
            /*
             * if (!healthy) { return
             * Response.status(503).entity("{\"status\":\"DOWN\"}").build(); }
             */

            // Creating a Workbook from an Excel file (.xls or .xlsx)
            Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));

            // Retrieving the number of sheets in the Workbook
            System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");

            /*
             * ============================================================= Iterating over
             * all the sheets in the workbook (Multiple ways)
             * =============================================================
             */

            // 1. You can obtain a sheetIterator and iterate over it
            Iterator<Sheet> sheetIterator = workbook.sheetIterator();
            System.out.println("Retrieving Sheets using Iterator");
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                System.out.println("=> " + sheet.getSheetName());
            }

            // 2. Or you can use a for-each loop
            System.out.println("Retrieving Sheets using for-each loop");
            for (Sheet sheet : workbook) {
                System.out.println("=> " + sheet.getSheetName());
            }

            // 3. Or you can use a Java 8 forEach with lambda
            System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
            workbook.forEach(sheet -> {
                System.out.println("=> " + sheet.getSheetName());
            });

            /*
             * ================================================================== Iterating
             * over all the rows and columns in a Sheet (Multiple ways)
             * ==================================================================
             */

            // Getting the Sheet at index zero
            Sheet sheet = workbook.getSheetAt(0);

            // Create a DataFormatter to format and get each cell's value as String
            DataFormatter dataFormatter = new DataFormatter();

            // 1. You can obtain a rowIterator and columnIterator and iterate over them
            System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
            Iterator<Row> rowIterator = sheet.rowIterator();
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();

                // Now let's iterate over the columns of the current row
                Iterator<Cell> cellIterator = row.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t");
                }
                System.out.println();
            }

            // 2. Or you can use a for-each loop to iterate over the rows and columns
            System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t");
                }
                System.out.println();
            }

            // 3. Or you can use Java 8 forEach loop with lambda
            System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
            sheet.forEach(row -> {
                row.forEach(cell -> {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t");
                });
                System.out.println();
            });

            // Closing the workbook
            workbook.close();

            XMLSlideShow powerpoint = new XMLSlideShow();
            XSLFSlide slide = powerpoint.createSlide();

            XSLFTable table = slide.createTable();
            table.setAnchor(new Rectangle(50, 50, 800, 800));

            int numColumns = 3;
            int numRows = 5;
            XSLFTableRow headerRow = table.addRow();
            headerRow.setHeight(50);
            // header
            for (int i = 0; i < numColumns; i++) {
                XSLFTableCell th = headerRow.addCell();
                XSLFTextParagraph p = th.addNewTextParagraph();
                p.setTextAlign(TextAlign.CENTER);
                XSLFTextRun r = p.addNewTextRun();
                r.setText("Header " + (i + 1));
                r.setFontSize(20.0);
                r.setFontColor(Color.white);
                th.setFillColor(new Color(79, 129, 189));
                table.setColumnWidth(i, 150);
            }

            // rows
            for (int rownum = 0; rownum < numRows; rownum++) {
                XSLFTableRow tr = table.addRow();
                tr.setHeight(50);
                // header
                for (int i = 0; i < numColumns; i++) {
                    XSLFTableCell cell = tr.addCell();
                    XSLFTextParagraph p = cell.addNewTextParagraph();
                    XSLFTextRun r = p.addNewTextRun();

                    // r.setText("Cell " + (i + 1));
                    r.setText("Cell " + (numRows * rownum + i + 1));
                    if (rownum % 2 == 0) {
                        cell.setFillColor(new Color(208, 216, 232));
                    } else {
                        cell.setFillColor(new Color(233, 247, 244));
                    }
                }
            }

            try {
                try (FileOutputStream out = new FileOutputStream("/home/relucio/microservicebuilder/poi/myFile.pptx")) {
                    try {
                        powerpoint.write(out);
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return Response.ok("{\"status\":\"UP\"}").build();
    }

}
