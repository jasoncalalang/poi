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
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.functions.Sumproduct;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
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
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
import org.apache.poi.util.POILogger;



public class HelloWorld {
//	public static final String SAMPLE_XLSX_FILE_PATH = "/home/ike/microservicebuilder/poi/spreadsheet.xlsx";
	private String xlsxPath = "/home/ike/Downloads/19Q1 Revenue Forecast (Week 8) Draft (When Editing - LOCK-OPEN-SAVE).xls";
	private String pptxTemplate = "/home/ike/Downloads/MOR-Template_all.pptx";
	private String pptxOutput = "/home/ike/Downloads/output.pptx";
	
	private String selectedMonth;

	private String deptCodes = "8E/7G/7Y/7H/9Y";
	private int numDepts = 5;

	private DataFormatter dataFormatter = new DataFormatter();
	private HSSFFormulaEvaluator evaluator;
	
	private int pptxChartOffset = 9999;
	
	private String garageName;
	private String garageNameReference = "'Garage & GEO Summary'!$F$5";
	
	// make money
	private int forecastColOffser = 9;
	private int backlogColOffset = 11;
	private String[] dept = new String[numDepts];
	private double[] deptForecast = new double[numDepts];
	private double[] deptBacklog = new double[numDepts];
	private double[] deptYTD = new double[numDepts];
    private double[][] deptMonthly = new double[numDepts][12];

	
	// revenue trends chart
	private double[] revenueTable = new double[12];
	private double[] quarterlyTable = new double[12];
	
	private HashMap<String, String> map = new HashMap<String, String>(){{
		put("8E", "Cloud Garage(8E)");
		put("7G", "Hybrid(7G)");
		put("7Y", "Analytics(7Y)");    
		put("9Y", "Blockchain(9Y)");    
		put("7H", "7H");    
		}};	
		

	
	
	private Workbook xlsworkbook;
	private Workbook pptworkbook;

	private SimpleDateFormat sdf = new SimpleDateFormat("MMMM");
	
	public static void main(String[] args) {
		HelloWorld o = new HelloWorld();
		o.setSelectedMonth("May");
		o.doit(args);
	}

		
	private void doit(String[] args) {
		
		// TODO Auto-generated method stub
		System.out.println("OK");
		System.out.println("we are here");
		try {

			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(getXlsxPath()));
//			HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File(getXlsxPath())));
			getGarageName(workbook);
			System.out.println("Garage : " + garageName);
			getMakeMoneyTable(workbook);
		    getRevenueTable(workbook);
		    getRevenueByDept(workbook);
			
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
				
				List <POIXMLDocumentPart> relations = slide.getRelations();
				for (int i = 0; i < relations.size(); i++) {
					POIXMLDocumentPart poixmlDocumentPart = (POIXMLDocumentPart) relations.get(i);
					System.out.println(poixmlDocumentPart);
					if (poixmlDocumentPart instanceof XSLFChart) {
						pptxChartOffset = i;
					}
					System.out.println(poixmlDocumentPart.getClass());
//					Name: /ppt/slideLayouts/slideLayout13.xml - Content Type: application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml
//					Name: /ppt/media/image2.tif - Content Type: image/tif
//					Name: /ppt/charts/chart1.xml - Content Type: application/vnd.openxmlformats-officedocument.drawingml.chart+xml
//					Name: /ppt/comments/comment1.xml - Content Type: application/vnd.openxmlformats-officedocument.presentationml.comments+xml
//					application/vnd.openxmlformats-officedocument.drawingml.chart+xml
				}

				System.out.println("\n");

				

				XSLFChart chart = (XSLFChart)relations.get(pptxChartOffset);
				XSSFWorkbook pptxworkbook = chart.getWorkbook();
				pptxworkbook.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);

				updateRevenueTrendsXLS(pptxworkbook, revenueTable);
				updateQuarterlyTrendsXLS(pptxworkbook, revenueTable);
				System.out.println("\n");

		        for (XSLFShape sh : slide.getShapes()) {
		            // name of the shape
		            String name = sh.getShapeName();
		            System.out.println("Shape name : " + name);

		            switch (name) {
					case "SlideTitle":
						// System.out.println(sh.getClass());
						((XSLFTextBox)sh).setText("IBM Cloud Garage " + garageName);
						// System.out.println("placeholder to change slide title");
						break;

					case "MakeMoneyTable":
	                	patchMakeMoneyTable((XSLFTable)sh);
	                	break;

					case "SpendMoneyTable":
	                	patchSpendMoneyTable((XSLFTable)sh);
	                	break;

					case "UtilizationTable":
	                	patchUtilizationTable((XSLFTable)sh);
	                	break;

					default:
						break;
					}
		            
//		            if (sh instanceof PlaceableShape) {
//		                java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
//		            }
//
//		            if (sh instanceof XSLFConnectorShape) {
//		                XSLFConnectorShape shape = (XSLFConnectorShape) sh;
//		                // System.out.println("is a Connector");
//		            } else if (sh instanceof XSLFTextShape) {
//		                XSLFTextShape shape = (XSLFTextShape) sh;
////		                System.out.println("is a Text SHape");
////		                System.out.println(shape.getText());
//		                
//		            } else if (sh instanceof XSLFPictureShape) {
//		                XSLFPictureShape shape = (XSLFPictureShape) sh;
////		                System.out.println("is a Picture Shape");
//		                
//		            } else if (sh instanceof XSLFTextBox) {
//		            	XSLFTextBox shape = (XSLFTextBox) sh;
////		                System.out.println("is a Text Box");
////		                System.out.println(shape.getText());
//		                
//		            } else if (sh instanceof XSLFTable) {
//		                		
//		            } else if (sh instanceof XSLFGraphicFrame) {
//		            	XSLFGraphicFrame shape = (XSLFGraphicFrame)sh;
////		            	System.out.println("is an Graphic Frame");
//		            	
//		            } else {
////		            	System.out.println("unknown instance");
////		            	System.out.println(sh);
//		            	
//		            }
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

	private void patchSpendMoneyTable(XSLFTable shape) {
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
		System.out.println("Patching Make Money Table");
		int numRows = shape.getNumberOfRows();
		int numCols = shape.getNumberOfColumns();

		for (int row = 0; row < dept.length; row++) {
			System.out.println(row + " " + numRows);
		}
		
		String fmtString = "#,##0.000";
		
		for (int row = 1; row < numRows; row++) {
			shape.getCell(row, 0).setText(dept[row - 1]);
			shape.getCell(row, 2).setText(formatDouble(deptBacklog[row - 1]/1000000, fmtString));
			shape.getCell(row, 3).setText(formatDouble(deptForecast[row - 1]/1000000, fmtString));
		}
		
	}
	
	private String formatDouble(double d, String formatStr) {
		return dataFormatter.formatRawCellContents(d, -1, formatStr);
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
		String[] targetRevenueMap = new String[] {
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
		
		System.out.println("Update Revenue Trends table");
		for (int i = 0; i < 12; i++) {
			Cell cell = getCell(wb, targetRevenueMap[i]);
			if (cell == null) {
				cell = createCell(wb, targetRevenueMap[i]);
			}
			setNumericCellValue(wb, targetRevenueMap[i], revenueTable[i]);
			System.out.println(cell);
		}

	}

	private void updateQuarterlyTrendsXLS(Workbook wb, double[] revenue) {
		System.out.println("Update Quarterly Trends table");
		
		String[] targetRevenueQtrMap = new String[] {
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
		
		int m = getMonthNumber(selectedMonth);
		int q = (int) (m/3);
		int[] qtr = new int [4];
		qtr[0] = 999;
		qtr[1] = 999;
		qtr[2] = 999;
		qtr[3] = 999;
		
		
		switch (q) {
		case 0:
			qtr[0] = m < q ? m : q;
			break;

		case 1:
			qtr[0] = 2;
			qtr[1] = m < q ? m : q;
			
			break;

		case 2:
			qtr[0] = 2;
			qtr[1] = 5;
			qtr[2] = m < q ? m : q;
			
			break;

		case 3:
			qtr[0] = 2;
			qtr[1] = 5;
			qtr[2] = 8;
			qtr[3] = m < q ? m : q;
			
			break;

		default:
			break;
		}
		
		for (int i = 0; i < 12; i++) {
			quarterlyTable[i] = 0.0;
		}
		
		for (int i = 0; i <= m; i++) {
			int qs = (int) (i/3);
			System.out.println("DOH");
			quarterlyTable[qtr[qs]] = quarterlyTable[qtr[qs]] + revenueTable[i];
		}
		
		for (int i = 0; i < 12; i++) {
			Cell cell = getCell(wb, targetRevenueQtrMap[i]);
			if (cell == null) {
				cell = createCell(wb, targetRevenueQtrMap[i]);
			}
			if ((i != qtr[0]) && (i != qtr[1]) && (i != qtr[2]) && (i != qtr[3])) {
				cell.setCellType(CellType.BLANK);
				continue;
			}
			setNumericCellValue(wb, targetRevenueQtrMap[i], quarterlyTable[i]);
//			System.out.println(cell);
			
		}
		
	}
	
	private void getRevenueTable(Workbook wb) {
		String[] revenueMap = new String[] {
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
		
		for (int i = 0; i < 12; i++) {
			if (i <= getMonthNumber(selectedMonth)) {
				revenueTable[i] = getNumericCellValue(wb, revenueMap[i]);
				System.out.println(revenueTable[i]);
			} else {
				revenueTable[i] = (double) 0;
			}
		}
		
	}

	private void getRevenueByDept(Workbook wb) {
		String[] map = new String[] {
				"'Garage & GEO Summary'!$F$30",
				"'Garage & GEO Summary'!$F$33",
				"'Garage & GEO Summary'!$F$36",
				"'Garage & GEO Summary'!$F$39",
				"'Garage & GEO Summary'!$F$42"
		};
		
		for (int i = 0; i < map.length; i++) {
			CellReference cellRef = new CellReference(map[i]);
			int pRow = cellRef.getRow();
			int pCol = cellRef.getCol();
			
			for (int j = 0; j < deptMonthly[i].length; j++) {
				deptMonthly[i][j] = 0.0;
				for (int k = 0; k < 3; k++) {
					CellReference cref = new CellReference("Garage & GEO Summary", pRow + k, pCol + j, true, true);
					deptMonthly[i][j] = deptMonthly[i][j] + getNumericCellValue(wb, cref.formatAsString());
				}
			}
		}
		
	}
	
	private double getNumericCellValue(Workbook wb, String cellStr) {
		Cell cell = getCell(wb, cellStr);
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
	
	private void getGarageName(Workbook wb) {
		String cellStr = "'Garage & GEO Summary'!$F$5";
		garageName = getStringCellValue(wb, cellStr);
	}
	
	private void getMakeMoneyTable(Workbook wb) {
		String cellStr = "'Garage & GEO Summary'!$D$14";

		System.out.println("Get Forecast and Backlog from worksheet");
		CellReference cellRef = new CellReference(cellStr);
		Sheet sheet = wb.getSheet(cellRef.getSheetName());
		for (int i = 0;i < numDepts; i++) {
			Row row = sheet.getRow(cellRef.getRow() + i);
			String deptCode = row.getCell(cellRef.getCol()).getStringCellValue();
			if (deptCodes.contains(deptCode.trim())) {
				dept[i] = deptCode;
				deptForecast[i] = row.getCell(cellRef.getCol() + forecastColOffser).getNumericCellValue();
				deptBacklog[i] = row.getCell(cellRef.getCol() + backlogColOffset).getNumericCellValue();
//				System.out.println(dataFormatter.formatCellValue(row.getCell(cellRef.getCol() + forecastColOffser)));
//				System.out.println(dataFormatter.formatCellValue(row.getCell(cellRef.getCol() + backlogColOffset)));
				System.out.println(deptCode);
				System.out.println("forecast = " + deptForecast[i]);
				System.out.println("backlog  = " + deptBacklog[i]);
			}
			
		}
		
	}


	public String getSelectedMonth() {
		return selectedMonth;
	}


	public void setSelectedMonth(String selectedMonth) {
		this.selectedMonth = selectedMonth;
	}
	
	private int getMonthNumber(String monthName) {
		int  i = 0;
		try {
			Date date = new SimpleDateFormat("MMMM").parse(monthName);
			Calendar cal = Calendar.getInstance();
			cal.setTime(date);
			i = cal.get(Calendar.MONTH);
		} catch (Exception e) {
		}
		
		return i;
	}
	
	
}
