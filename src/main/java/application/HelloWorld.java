package application;

import java.awt.Color;
import java.awt.Rectangle;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData.Series;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFChart;
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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTChartImpl;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.sl.usermodel.SlideShowFactory;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFGraphicFrame;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.util.IOUtils;



public class HelloWorld {
//	public static final String SAMPLE_XLSX_FILE_PATH = "/home/ike/microservicebuilder/poi/spreadsheet.xlsx";
	public static final String SAMPLE_XLSX_FILE_PATH = "/home/ike/Downloads/19Q1 Revenue Forecast (Week 8) Draft (When Editing - LOCK-OPEN-SAVE).xls";
	public static final String SAMPLE_PPTX_FILE_PATH = "/home/ike/Downloads/MOR-POI-Template.pptx";
	public static final String OUTPUT_PPTX_FILE_PATH = "/tmp/output.pptx";

	private DataFormatter dataFormatter = new DataFormatter();
	
	private double[] revenueTable = new double[13];
	private Workbook xlsworkbook;
	private Workbook pptworkbook;
	
	public static void main(String[] args) {
		HelloWorld o = new HelloWorld();
		o.doit(args);
	}

		
	private void doit(String[] args) {
		
		// TODO Auto-generated method stub
		System.out.println("OK");
		System.out.println("we are here");
		try {

			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(SAMPLE_XLSX_FILE_PATH));
			// DataFormatter dataFormatter = new DataFormatter();

			// Retrieving the number of sheets in the Workbook
			System.out.println("Workbook has " + workbook.getNumberOfSheets() + " Sheets : ");
			
			List<? extends Name> names = workbook.getAllNames();
			Iterator<? extends Name> it = names.iterator();
			while (it.hasNext()) {
				Name type = (Name) it.next();
				System.out.println(type.getNameName());
				// System.out.println(type.getRefersToFormula());
			}

//			Name range1 = workbook.getName("range1");
//			// retrieve the cell at the named range and test its contents
//		    AreaReference aref = new AreaReference(range1.getRefersToFormula(),SpreadsheetVersion.EXCEL2007);
//		    CellReference[] crefs = aref.getAllReferencedCells();
//		    for (int i=0; i<crefs.length; i++) {
//		        Sheet s = workbook.getSheet(crefs[i].getSheetName());
//		        Row r = s.getRow(crefs[i].getRow());
//		        Cell c = r.getCell(crefs[i].getCol());
//		        System.out.println(dataFormatter.formatCellValue(c));
//		        // extract the cell contents based on cell type etc.
//		    }
			
			Iterator<Sheet> sheetIterator = workbook.sheetIterator();
			System.out.println("Retrieving Sheets using Iterator");
			while (sheetIterator.hasNext()) {
				Sheet sheet = sheetIterator.next();
				System.out.println("=> " + sheet.getSheetName());
			}

			Sheet sheet = workbook.getSheetAt(0);
			workbook.close();
			
			
			ZipSecureFile.setMinInflateRatio(-1.0d);
			XMLSlideShow powerpoint = new XMLSlideShow(new FileInputStream(SAMPLE_PPTX_FILE_PATH));

			List<PackagePart> pp = powerpoint.getAllEmbeddedParts();
			Iterator<PackagePart> ppi = pp.iterator();
			while (ppi.hasNext()) {
				PackagePart object = (PackagePart) ppi.next();
				System.out.println("Package name = " + object.getPartName());
			}
			
			System.out.println("\n");
			
			for (XSLFSlide slide : powerpoint.getSlides()) {
				System.out.println("Slide : " + slide.getSlideName());
				XSLFTextShape [] ts = slide.getPlaceholders();
				System.out.println("Placeholders");
				for (int i = 0; i < ts.length; i++) {
					System.out.println(ts[i].getText());
					
				}
				System.out.println("\n");
				
				System.out.println("Relations");
				List <POIXMLDocumentPart> relations = slide.getRelations();
				for (Iterator<POIXMLDocumentPart> iterator = relations.iterator(); iterator.hasNext();) {
					POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart) iterator.next();
					System.out.println(poixmlDocumentPart);
				}
				System.out.println(relations.get(2).getPackagePart().getContentType());
			
				System.out.println("\n");
				
				System.out.println("Chart");

				XSLFChart chart = (XSLFChart)relations.get(2);
				List<XDDFChartData> chartdata = chart.getChartSeries();
				for (Iterator iterator = chartdata.iterator(); iterator.hasNext();) {
					XDDFChartData xddfChartData = (XDDFChartData) iterator.next();
					List<Series> series = xddfChartData.getSeries();
					for (Iterator iterator2 = series.iterator(); iterator2.hasNext();) {
						Series series2 = (Series) iterator2.next();
						series2.getValuesData().getDataRangeReference();
					}
				}
				
				chart.getCTChart().save(new File("/tmp/ctchart"));
				XSSFWorkbook pptxworkbook = chart.getWorkbook();
				pptxworkbook.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);

				revenueTable[1]  = 10000;
				revenueTable[2]  = 20000;
				revenueTable[3]  = 30000;
				revenueTable[4]  = 40000;
				revenueTable[5]  = 50000;
				revenueTable[6]  = 60000;
				revenueTable[7]  = 70000;
				revenueTable[8]  = 80000;
				revenueTable[9]  = 90000;
				revenueTable[10] = 100000;
				revenueTable[11] = 110000;
				revenueTable[12] = 120000;
				
				updateRevenueTrendsXLS(pptxworkbook.getSheetAt(0), revenueTable);
				
				System.out.println("\n");
				

		        for (XSLFShape sh : slide.getShapes()) {
		            // name of the shape
		            String name = sh.getShapeName();
		            System.out.println("Shape name : " + name);

		            if (sh instanceof PlaceableShape) {
		                java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
		            }

		            if (sh instanceof XSLFConnectorShape) {
		                XSLFConnectorShape line = (XSLFConnectorShape) sh;
		                // System.out.println("is a Connector");
		            } else if (sh instanceof XSLFTextShape) {
		                XSLFTextShape shape = (XSLFTextShape) sh;
//		                System.out.println("is a Text SHape");
//		                System.out.println(shape.getText());
		                
		            } else if (sh instanceof XSLFPictureShape) {
		                XSLFPictureShape shape = (XSLFPictureShape) sh;
//		                System.out.println("is a Picture Shape");
		                
		            } else if (sh instanceof XSLFTextBox) {
		            	XSLFTextBox shape = (XSLFTextBox) sh;
//		                System.out.println("is a Text Box");
//		                System.out.println(shape.getText());
		                
		            } else if (sh instanceof XSLFTable) {
		            	XSLFTable shape = (XSLFTable) sh;
		                System.out.println("is a Table");
		                if (name.equals("MakeMoneyTable")) {
		                	patchMakeMoneyTable(shape);
		                	
		                } else if (name.equals("SpendMoneyTable")) {
		                	patchSendMoneyTable(shape);
		                	
		                } else if (name.equals("UtilizationTable")) {
		                	patchUtilizationTable(shape);		                	
		                }
		                		
		            } else if (sh instanceof XSLFGraphicFrame) {
		            	XSLFGraphicFrame shape = (XSLFGraphicFrame)sh;
//		            	System.out.println("is an Graphic Frame");
		            	
		            } else {
//		            	System.out.println("unknown instance");
//		            	System.out.println(sh);
		            	
		            }
		            System.out.println("\n\n");
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
			try {
				try (FileOutputStream out = new FileOutputStream(OUTPUT_PPTX_FILE_PATH)) {
					try {
						powerpoint.write(out);
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
			powerpoint.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void patchUtilizationTable(XSLFTable shape) {
		// TODO Auto-generated method stub
		System.out.println("Patching Utilization Table");
		int numRows = shape.getNumberOfRows();
		int numCols = shape.getNumberOfColumns();
		
		for (int row = 1; row < numRows; row++) {
			for (int col = 1; col < numCols; col++) {
				shape.getCell(row, col).setText("U");
			}
		}
		
	}

	private void patchSendMoneyTable(XSLFTable shape) {
		// TODO Auto-generated method stub
		System.out.println("Patching Send Money Table");
		int numRows = shape.getNumberOfRows();
		int numCols = shape.getNumberOfColumns();
		
		for (int row = 1; row < numRows; row++) {
			for (int col = 1; col < numCols; col++) {
				shape.getCell(row, col).setText("S");
			}
		}
		
	}

	private void patchMakeMoneyTable(XSLFTable shape) {
		// TODO Auto-generated method stub
		System.out.println("Pathcing Make Money Table");
		int numRows = shape.getNumberOfRows();
		int numCols = shape.getNumberOfColumns();
		
		for (int row = 1; row < numRows; row++) {
			for (int col = 1; col < numCols; col++) {
				shape.getCell(row, col).setText("M");
			}
		}
		
	}
	
	private void dumpWorkbookSheet(Sheet sheet) {
		System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
		Iterator<Row> rowIterator = sheet.rowIterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			// Now let's iterate over the columns of the current row
			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				System.out.println(cell.getCellType().name());
				String cellValue = dataFormatter.formatCellValue(cell);
				System.out.print(cellValue + "\t");
			}
			System.out.println();
		}
		
	}
	
	private void updateRevenueTrendsXLS(Sheet sheet, double[] revenue) {
		double q1 = revenue[1] + revenue[2] + revenue[3];
		double q2 = revenue[4] + revenue[5] + revenue[6];
		double q3 = revenue[7] + revenue[8] + revenue[9];
		double q4 = revenue[10] + revenue[11] + revenue[12];

		for (int rowNum = 1; rowNum < 13; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row.getCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK) == null) {
				row.createCell(2);
			}
			System.out.println(row.getCell(2, MissingCellPolicy.RETURN_NULL_AND_BLANK));
//			row.getCell(2).setCellValue(revenue[rowNum]);
			setCellValue(row, 2, revenue[rowNum]);
		}
		
		setCellValue(sheet.getRow(3), 1, q1);
		setCellValue(sheet.getRow(6), 1, q2);
		setCellValue(sheet.getRow(9), 1, q3);
		setCellValue(sheet.getRow(12), 1, q4);

	}
	
	private void setCellValue(Row row, int cellidx, double val) {
		if (row.getCell(cellidx, MissingCellPolicy.RETURN_NULL_AND_BLANK) == null) {
			row.createCell(cellidx);
		}
		row.getCell(cellidx).setCellValue(val);
	}

}
