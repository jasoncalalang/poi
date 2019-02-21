package application;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTableRow;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFTable;


public class HelloWorld {
	public static final String SAMPLE_XLSX_FILE_PATH = "/home/relucio/microservicebuilder/poi/spreadsheet.xlsx";
	public static final String SAMPLE_PPTX_FILE_PATH = "/home/relucio/Downloads/MOR-Template.pptx";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		System.out.println("OK");
		System.out.println("we are here");
		try {
			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
			DataFormatter dataFormatter = new DataFormatter();

			// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
			
			List<? extends Name> names = workbook.getAllNames();
			Iterator it = names.iterator();
			while (it.hasNext()) {
				XSSFName type = (XSSFName) it.next();
				System.out.println(type.getNameName());
				System.out.println(type.getRefersToFormula());
			}

			Name range1 = workbook.getName("range1");
			// retrieve the cell at the named range and test its contents
		    AreaReference aref = new AreaReference(range1.getRefersToFormula(),SpreadsheetVersion.EXCEL2007);
		    CellReference[] crefs = aref.getAllReferencedCells();
		    for (int i=0; i<crefs.length; i++) {
		        Sheet s = workbook.getSheet(crefs[i].getSheetName());
		        Row r = s.getRow(crefs[i].getRow());
		        Cell c = r.getCell(crefs[i].getCol());
		        System.out.println(dataFormatter.formatCellValue(c));
		        // extract the cell contents based on cell type etc.
		    }
			
			
			
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

//			// 2. Or you can use a for-each loop
//			System.out.println("Retrieving Sheets using for-each loop");
//			for (Sheet sheet : workbook) {
//				System.out.println("=> " + sheet.getSheetName());
//			}
//
//			// 3. Or you can use a Java 8 forEach with lambda
//			System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
//			workbook.forEach(sheet -> {
//				System.out.println("=> " + sheet.getSheetName());
//			});

			/*
			 * ================================================================== Iterating
			 * over all the rows and columns in a Sheet (Multiple ways)
			 * ==================================================================
			 */

			// Getting the Sheet at index zero
			Sheet sheet = workbook.getSheetAt(0);

			// Create a DataFormatter to format and get each cell's value as String

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

//			// 2. Or you can use a for-each loop to iterate over the rows and columns
//			System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
//			for (Row row : sheet) {
//				for (Cell cell : row) {
//					String cellValue = dataFormatter.formatCellValue(cell);
//					System.out.print(cellValue + "\t");
//				}
//				System.out.println();
//			}
//
//			// 3. Or you can use Java 8 forEach loop with lambda
//			System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
//			sheet.forEach(row -> {
//				row.forEach(cell -> {
//					String cellValue = dataFormatter.formatCellValue(cell);
//					System.out.print(cellValue + "\t");
//				});
//				System.out.println();
//			});

			// Closing the workbook
			workbook.close();
			ZipSecureFile.setMinInflateRatio(-1.0d);
			XMLSlideShow powerpoint = new XMLSlideShow(new FileInputStream(SAMPLE_PPTX_FILE_PATH));

//			System.out.println("Available slide layouts:");
//
//			//getting the list of all slide masters
//			for(XSLFSlideMaster master : powerpoint.getSlideMasters()) {
//				//getting the list of the layouts in each slide master
//				for(XSLFSlideLayout layout : master.getSlideLayouts()) {
//
//					//getting the list of available slides
//					System.out.println(layout.getType());
//				} 
//			}            
//
			
			for (XSLFSlide slide : powerpoint.getSlides()) {
				System.out.println("Slide : " + slide.getSlideName());
		        for (XSLFShape sh : slide.getShapes()) {
		            // name of the shape
		            String name = sh.getShapeName();
		            System.out.println("Shape name : " + name);
		            // System.out.println(sh);
		            // shapes's anchor which defines the position of this shape in the slide
		            if (sh instanceof PlaceableShape) {
		                java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
		            }

		            if (sh instanceof XSLFConnectorShape) {
		                XSLFConnectorShape line = (XSLFConnectorShape) sh;
		                // work with Line
		            } else if (sh instanceof XSLFTextShape) {
		                XSLFTextShape shape = (XSLFTextShape) sh;
		                // work with a shape that can hold text
		            } else if (sh instanceof XSLFPictureShape) {
		                XSLFPictureShape shape = (XSLFPictureShape) sh;
		                // work with Picture
		            } else if (sh instanceof XSLFTextBox) {
		            	XSLFTextBox shape = (XSLFTextBox) sh;
		            } else if (sh instanceof XSLFTable) {
		            	XSLFTable shape = (XSLFTable) sh;
		            } else {
		            	System.out.println(sh);
		            }
		        }
		    }			
			
			
//			XSLFSlideMaster defaultMaster = powerpoint.getSlideMasters().get(0);  
//			XSLFSlideLayout tc = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);  
//
//			XSLFSlide slide = powerpoint.createSlide(tc);
//
//			XSLFTable table = slide.createTable();
//			table.setAnchor(new Rectangle(50, 50, 800, 800));
//
//			int numColumns = 3;
//			int numRows = 5;
//			XSLFTableRow headerRow = table.addRow();
//			headerRow.setHeight(50);
//			// header
//			for (int i = 0; i < numColumns; i++) {
//				XSLFTableCell th = headerRow.addCell();
//				XSLFTextParagraph p = th.addNewTextParagraph();
//				p.setTextAlign(TextAlign.CENTER);
//				XSLFTextRun r = p.addNewTextRun();
//				r.setText("Header " + (i + 1));
//				r.setFontSize(20.0);
//				r.setFontColor(Color.white);
//				th.setFillColor(new Color(79, 129, 189));
//				table.setColumnWidth(i, 150);
//			}
//
//			// rows
//			for (int rownum = 0; rownum < numRows; rownum++) {
//				XSLFTableRow tr = table.addRow();
//				tr.setHeight(50);
//				// header
//				for (int i = 0; i < numColumns; i++) {
//					XSLFTableCell cell = tr.addCell();
//					XSLFTextParagraph p = cell.addNewTextParagraph();
//					XSLFTextRun r = p.addNewTextRun();
//
//					// r.setText("Cell " + (i + 1));
//					r.setText("Cell " + (rownum*numColumns + i));
//					if (rownum % 2 == 0) {
//						cell.setFillColor(new Color(208, 216, 232));
//					} else {
//						cell.setFillColor(new Color(233, 247, 244));
//					}
//				}
//			}
//
//			try {
//				try (FileOutputStream out = new FileOutputStream("/home/ike/microservicebuilder/poi/myFile.pptx")) {
//					try {
//						powerpoint.write(out);
//						powerpoint.close();
//					} catch (IOException e) {
//						e.printStackTrace();
//					}
//				}
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
