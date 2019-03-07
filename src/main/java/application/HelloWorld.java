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
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
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
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.chart.impl.CTChartImpl;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
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
	private String xlsxPath = "/home/relucio/Downloads/19Q1 Revenue Forecast (Week 8) Draft (When Editing - LOCK-OPEN-SAVE).xls";
	private String pptxTemplate = "/home/relucio/Downloads/MOR-POI-Template.pptx";
	private String pptxOutput = "/tmp/output.pptx";

	private DataFormatter dataFormatter = new DataFormatter();
	
	private int pptxChartOffset = 0; 
	
	private double[] revenueTable = new double[13];
	private String[] revenueMap = new String[] {
			"",
			"'Garage & GEO Summary'!$F$46",
			"'Garage & GEO Summary'!$G$46",
			"'Garage & GEO Summary'!$H$46",
			"'Garage & GEO Summary'!$I$46",
			"'Garage & GEO Summary'!$J$46",
			"'Garage & GEO Summary'!$K$46",
			"'Garage & GEO Summary'!$L$46",
			"'Garage & GEO Summary'!$M$46",
			"'Garage & GEO Summary'!$N$46",
			"'Garage & GEO Summary'!$O$46",
			"'Garage & GEO Summary'!$P$46",
			"'Garage & GEO Summary'!$Q$46"
	};
	
	private String[] targetRevenueMap = new String[] {
			"",
			"'Sheet1'!$C$2",
			"'Sheet1'!$C$3",
			"'Sheet1'!$C$4",
			"'Sheet1'!$C$5",
			"'Sheet1'!$C$6",
			"'Sheet1'!$C$7",
			"'Sheet1'!$C$8",
			"'Sheet1'!$C$9",
			"'Sheet1'!$C$10",
			"'Sheet1'!$C$11",
			"'Sheet1'!$C$12",
			"'Sheet1'!$C$13"
	};

	private String[] targetRevenueQtrMap = new String[] {
			"",
			"'Sheet1'!$B$2",
			"'Sheet1'!$B$3",
			"'Sheet1'!$B$4",
			"'Sheet1'!$B$5",
			"'Sheet1'!$B$6",
			"'Sheet1'!$B$7",
			"'Sheet1'!$B$8",
			"'Sheet1'!$B$9",
			"'Sheet1'!$B$10",
			"'Sheet1'!$B$11",
			"'Sheet1'!$B$12",
			"'Sheet1'!$B$13"
	};
	
	
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
			Workbook workbook = WorkbookFactory.create(new File(getXlsxPath()));
			
			// retrieveRevenueTable(workbook);
			// Sheet sheet = workbook.getSheet("Garage & GEO Summary");
			
			System.out.println(getStringCellValue(workbook,"'Garage & GEO Summary'!$F$5"));
			System.out.println("===================" + getNumericCellValue(workbook, "'Garage & GEO Summary'!$F$14"));
			setStringCellValue(workbook, "'Garage & GEO Summary'!$F$5", "Singapore");
			System.out.println(getStringCellValue(workbook,"'Garage & GEO Summary'!$F$5"));
			// HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			for (Sheet sheet : workbook) {
			    for (Row r : sheet) {
			        for (Cell c : r) {
//			            if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
				        if (c.getCellType().equals("FORMULA")) {
			                evaluator.evaluateFormulaCell(c);
			            }
			        }
			    }
			}			
			HSSFFormulaEvaluator.evaluateAllFormulaCells(workbook);
			System.out.println("===================" + getNumericCellValue(workbook, "'Garage & GEO Summary'!$F$14"));
			workbook.close();
			
			
			ZipSecureFile.setMinInflateRatio(-1.0d);
			XMLSlideShow powerpoint = new XMLSlideShow(new FileInputStream(getPptxTemplate()));

//			List<PackagePart> pp = powerpoint.getAllEmbeddedParts();
//			Iterator<PackagePart> ppi = pp.iterator();
//			while (ppi.hasNext()) {
//				PackagePart object = (PackagePart) ppi.next();
//				System.out.println("Package name = " + object.getPartName());
//			}
			
			
			for (XSLFSlide slide : powerpoint.getSlides()) {
//				System.out.println("Slide : " + slide.getSlideName());
//				XSLFTextShape [] ts = slide.getPlaceholders();
//				System.out.println("Placeholders");
//				for (int i = 0; i < ts.length; i++) {
//					System.out.println(ts[i].getText());
//					
//				}
//				System.out.println("\n");
				
				List <POIXMLDocumentPart> relations = slide.getRelations();

//				for (Iterator<POIXMLDocumentPart> iterator = relations.iterator(); iterator.hasNext();) {
//					POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart) iterator.next();
//					System.out.println(poixmlDocumentPart);
//				}
//				System.out.println("\n");

//				Relations -- we want the 3rd one (2)
//				Name: /ppt/slideLayouts/slideLayout13.xml - Content Type: application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml
//				Name: /ppt/media/image2.tif - Content Type: image/tif
//				Name: /ppt/charts/chart1.xml - Content Type: application/vnd.openxmlformats-officedocument.drawingml.chart+xml
//				Name: /ppt/comments/comment1.xml - Content Type: application/vnd.openxmlformats-officedocument.presentationml.comments+xml
//				application/vnd.openxmlformats-officedocument.drawingml.chart+xml
				
				pptxChartOffset = 2;

				XSLFChart chart = (XSLFChart)relations.get(pptxChartOffset);
//				List<XDDFChartData> chartdata = chart.getChartSeries();
//				for (Iterator iterator = chartdata.iterator(); iterator.hasNext();) {
//					XDDFChartData xddfChartData = (XDDFChartData) iterator.next();
//					List<Series> series = xddfChartData.getSeries();
//					for (Iterator iterator2 = series.iterator(); iterator2.hasNext();) {
//						Series series2 = (Series) iterator2.next();
//						series2.getValuesData().getDataRangeReference();
//					}
//				}
				
				XSSFWorkbook pptxworkbook = chart.getWorkbook();
				pptxworkbook.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);
				updateRevenueTrendsXLS(pptxworkbook, revenueTable);
				System.out.println("\n");

		        for (XSLFShape sh : slide.getShapes()) {
		            // name of the shape
		            String name = sh.getShapeName();
		            System.out.println("Shape name : " + name);

		            if (sh instanceof PlaceableShape) {
		                java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
		            }

		            if (sh instanceof XSLFConnectorShape) {
		                XSLFConnectorShape shape = (XSLFConnectorShape) sh;
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
				try (FileOutputStream out = new FileOutputStream(getPptxOutput())) {
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
	
	private void updateRevenueTrendsXLS(Workbook wb, double[] revenue) {
		Sheet sheet = wb.getSheet("Sheet1");
		double q1 = revenue[1] + revenue[2] + revenue[3];
		double q2 = revenue[4] + revenue[5] + revenue[6];
		double q3 = revenue[7] + revenue[8] + revenue[9];
		double q4 = revenue[10] + revenue[11] + revenue[12];

		for (int i = 1; i < 13; i++) {
			Cell cell = getCell(wb, targetRevenueMap[i]);
			if (cell == null) {
				cell = createCell(wb, targetRevenueMap[i]);
			}
			setNumericCellValue(wb, targetRevenueMap[i], revenueTable[i]);
			System.out.println(cell);
		}
		
//		setCellValue(sheet.getRow(3), 1, q1);
//		setCellValue(sheet.getRow(6), 1, q2);
//		setCellValue(sheet.getRow(9), 1, q3);
//		setCellValue(sheet.getRow(12), 1, q4);

	}
	
	private void setCellValue(Row row, int cellidx, double val) {
		if (row.getCell(cellidx, MissingCellPolicy.RETURN_NULL_AND_BLANK) == null) {
			row.createCell(cellidx);
		}
		row.getCell(cellidx).setCellValue(val);
	}
	
	private void retrieveRevenueTable(Workbook wb) {
		
		for (int i = 1; i < 13; i++) {
			revenueTable[i] = getNumericCellValue(wb, revenueMap[i]);
		}
		
	}
	
	private double getNumericCellValue(Workbook wb, String cellStr) {
		Cell cell = getCell(wb, cellStr);
		System.out.println(cell.getNumericCellValue());
		return cell.getNumericCellValue();
	}

	private String getStringCellValue(Workbook wb, String cellStr) {
		Cell cell = getCell(wb, cellStr);
		return(cell.getStringCellValue());
	}


	private void setNumericCellValue(Workbook wb, String cellStr, double value) {
		Cell cell = getCell(wb, cellStr);
		cell.setCellValue(value);
	}

	private void setStringCellValue(Workbook wb, String cellStr, String value) {
		Cell cell = getCell(wb, cellStr);
		cell.setCellValue(value);
	}
	
	private Cell getCell(Workbook wb, String cellStr) {
		CellReference c = new CellReference(cellStr);
		Sheet sheet = wb.getSheet(c.getSheetName());
		Cell cell = sheet.getRow(c.getRow()).getCell(c.getCol());
		return cell;
	}

	private Cell createCell(Workbook wb, String cellStr) {
		CellReference c = new CellReference(cellStr);
		Sheet sheet = wb.getSheet(c.getSheetName());
		Cell cell = sheet.getRow(c.getRow()).getCell(c.getCol());
		if (cell == null) {
			cell = sheet.getRow(c.getRow()).createCell(c.getCol());
		}
		return cell;
	}


	private String getXlsxPath() {
		return xlsxPath;
	}


	private void setXlsxPath(String xlsxPath) {
		this.xlsxPath = xlsxPath;
	}


	private String getPptxTemplate() {
		return pptxTemplate;
	}


	private void setPptxTemplate(String pptxTemplate) {
		this.pptxTemplate = pptxTemplate;
	}


	private String getPptxOutput() {
		return pptxOutput;
	}


	private void setPptxOutput(String pptxOutput) {
		this.pptxOutput = pptxOutput;
	}
	
	
}
