package application;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFChart;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.util.ZipSecureFile;

//  THIS WILL MAKE AN EXCELLENT EXERCISE IN REFACTORING

public class HelloWorld {
//	public static final String SAMPLE_XLSX_FILE_PATH = "/home/relucio/microservicebuilder/poi/spreadsheet.xlsx";
	// private String xlsxPath = "/home/relucio/Downloads/19Q1 Revenue Forecast (Week 8) Draft (When Editing - LOCK-OPEN-SAVE).xls";
	// private String xlsxPath = "/home/relucio/Downloads/19Q1 Revenue Forecast (Week 8) Draft (When Editing - LOCK-OPEN-SAVE).xls";
	private String xlsxPath = "/home/relucio/Downloads/19Q2 Revenue Forecast (Week 3) Draft (When Editing - LOCK - OPEN - SAVE).xls";
	
	private String utilizationXlsxPath = "/home/relucio/Downloads/191Q Utilization Report weekending 040519.xlsx";
	private String pptxTemplate = "/home/relucio/Downloads/MOR-Template_all_copy.pptx";
	private String pptxOutput = "/home/relucio/Downloads/output.pptx";
	
	private String selectedMonth = "Jan";
	private String selectedYear = "2019";

	private String deptCodes = "8E/7G/7Y/7H/9Y";
	private int numDepts = 5;
	private DataFormatter dataFormatter = new DataFormatter();
	
	private int pptxChartOffset = 9999;
	
	private String garageName;
	
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

	
	// utilization table
	private String[] qtrMonth = new String[3];
	private double[] monthTotal = new double[3];
	private double[] monthBillable = new double[3];
	private double[] monthCustomerFacing = new double[3];
	private double totalHours = 0.0;
	private double billableHours = 0.0;
	private double customerFacingHours = 0.0;
	
	@SuppressWarnings("serial")
	private HashMap<String, String> deptNames = new HashMap<String, String>(){{
		put("8E", "Cloud Garage(8E)");
		put("7G", "Hybrid(7G)");
		put("7Y", "Analytics(7Y)");    
		put("9Y", "Blockchain(9Y)");    
		put("7H", "7H");    
		}};	
		

	
	

	
	public static void main(String[] args) {
		HelloWorld o = new HelloWorld();
		o.setSelectedMonth("May");
		o.setSelectedYear("2019");
		o.doit(args);
	}

		
	private void doit(String[] args) {
		
		try {

			// Creating a Workbook from an Excel file (.xls or .xlsx)
			Workbook workbook = WorkbookFactory.create(new File(getXlsxPath()));
			Workbook utilizationWorkbook = WorkbookFactory.create(new File(getUtilizationXlsxPath()));
			XMLSlideShow powerpoint = process(workbook, utilizationWorkbook);
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

	
	public XMLSlideShow process(Workbook workbook, Workbook utilizationWorkbook) {

		XMLSlideShow powerpoint = null;
		try {

			getGarageName(workbook);

			System.out.println("Garage : " + garageName);

			getMakeMoneyTable(workbook);
			getRevenueTable(workbook);
			getRevenueByDept(workbook);
			getUtilization(utilizationWorkbook);
			workbook.close();
			utilizationWorkbook.close();


			ZipSecureFile.setMinInflateRatio(-1.0d);
			powerpoint = new XMLSlideShow(new FileInputStream(getPptxTemplate()));

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

				XSLFChart chart = (XSLFChart)relations.get(pptxChartOffset);
				XSSFWorkbook pptxworkbook = chart.getWorkbook();
				pptxworkbook.setMissingCellPolicy(MissingCellPolicy.RETURN_NULL_AND_BLANK);

				updateRevenueTrendsXLS(pptxworkbook, revenueTable);
				updateQuarterlyTrendsXLS(pptxworkbook, revenueTable);
				
				for (XSLFShape sh : slide.getShapes()) {
					String name = sh.getShapeName();

					switch (name) {
					case "SlideTitle":
						((XSLFTextBox)sh).setText("IBM Cloud Garage " + garageName);
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
						
					case "ResourceUtilizationTitle":
						pathUtilizationTitle((XSLFTextBox)sh);
						break;
						
					case "IBMCloudGarageFooter":
						System.out.println("FOUND FOOTER");
						System.out.println(sh);
						break;

					default:
						break;
					}

				}
			}			

		} catch (Exception e) {
			e.printStackTrace();
		}
		return powerpoint;
	}
	
	private void pathUtilizationTitle(XSLFTextBox sh) {
		// TODO Auto-generated method stub
		System.out.println(sh.getText());
		System.out.println(getQuarter());
		System.out.println(sh.getTextParagraphs().size());
		System.out.println(sh.getTextParagraphs().get(1).getText());
		System.out.println(sh.getTextParagraphs().get(1).getTextRuns().get(0).getRawText());
		sh.getTextParagraphs().get(1).getTextRuns().get(0).setText("Average " + getQuarter() + "/" + getSelectedYear() +"*");
		
	}


	private void getUtilization(Workbook utilizationWorkbook) {
		// TODO Auto-generated method stub
		XSSFSheet sheet = ((XSSFWorkbook)utilizationWorkbook).getSheet("Adhoc Export");
		System.out.println("First row = " + sheet.getFirstRowNum());
		System.out.println("Top row = " + sheet.getTopRow());
		
		short totalCol = (new CellReference("AD1")).getCol();
		short billableCol = (new CellReference("AF1")).getCol();
		short customerFacingCol = (new CellReference("AH1").getCol());
		short monthCol = (new CellReference("U1").getCol());
		
		double t;
		double b;
		double c;
		
		monthTotal[0] = 0.0;
		monthTotal[1] = 0.0;
		monthTotal[2] = 0.0;
		
		monthBillable[0] = 0.0;
		monthBillable[1] = 0.0;
		monthBillable[2] = 0.0;
		
		monthCustomerFacing[0] = 0.0;
		monthCustomerFacing[1] = 0.0;
		monthCustomerFacing[2] = 0.0;
		
		qtrMonth = getQuarterMonths();
		
		String m = "";
		
		Iterator<Row> i = sheet.iterator();
		while (i.hasNext()) {
			Row row = (Row) i.next();
			if (row.getRowNum() == 0) continue;
			Cell cell = row.getCell(4);
			if ((cell != null) && (!row.getZeroHeight())) {
				t = row.getCell(totalCol).getNumericCellValue();
				b = row.getCell(billableCol).getNumericCellValue();
				c = row.getCell(customerFacingCol).getNumericCellValue();
				m = row.getCell(monthCol).getStringCellValue();
				// System.out.println("Month = " + m);
				
				if (qtrMonth[0].equals(m)) {
					monthTotal[0] = monthTotal[0] + t;
					monthBillable[0] = monthBillable[0] + b;
					monthCustomerFacing[0] = monthCustomerFacing[0] + c;
				} else if (qtrMonth[1].equals(m)) {
					monthTotal[1] = monthTotal[1] + t;
					monthBillable[1] = monthBillable[1] + b;
					monthCustomerFacing[1] = monthCustomerFacing[1] + c;
				} else if (qtrMonth[2].equals(m)) {
					monthTotal[2] = monthTotal[2] + t;
					monthBillable[2] = monthBillable[2] + b;
					monthCustomerFacing[2] = monthCustomerFacing[2] + c;
				} else {

				}
				
				
				totalHours = totalHours + t;
				billableHours = billableHours + b;
				customerFacingHours = customerFacingHours + c;
			}
		}
		
		System.out.println(sheet.getActiveCell().formatAsString());
		
	}

	private String getQuarter() {
		return "Q" + ((int)getMonthNumber(selectedMonth)/3 + 1);
	}
	private String[] getQuarterMonths() {
		String[] qm = new String[3];
		int m = getMonthNumber(selectedMonth);
		int q = (int)m/3;
		switch (q) {
		case 0:
			qm[0] = "JAN";
			qm[1] = "FEB";
			qm[2] = "MAR";
			break;

		case 1:
			qm[0] = "APR";
			qm[1] = "MAY";
			qm[2] = "JUN";
			break;

		case 2:
			qm[0] = "JUL";
			qm[1] = "AUG";
			qm[2] = "SEP";
			break;

		case 3:
			qm[0] = "OCT";
			qm[1] = "NOV";
			qm[2] = "DEC";
			break;

		default:
			break;
		}
		return qm;
	}


	private void patchUtilizationTable(XSLFTable shape) {
		// TODO Auto-generated method stub
		System.out.println("Patching Utilization Table");
		int numRows = shape.getNumberOfRows();
		int numCols = shape.getNumberOfColumns();
		

		shape.getCell(0,1).getTextParagraphs().get(0).getTextRuns().get(0).setText(qtrMonth[0]);
		shape.getCell(0,2).getTextParagraphs().get(0).getTextRuns().get(0).setText(qtrMonth[1]);
		shape.getCell(0,3).getTextParagraphs().get(0).getTextRuns().get(0).setText(qtrMonth[2]);
		
		String fmtString = "0.00%";
		
		// billable row
		shape.getCell(1, 1).setText(formatDoubleAsPercentage(monthBillable[0],monthTotal[0]));
		shape.getCell(1, 2).setText(formatDoubleAsPercentage(monthBillable[1],monthTotal[1]));
		shape.getCell(1, 3).setText(formatDoubleAsPercentage(monthBillable[2],monthTotal[2]));
		shape.getCell(1, 4).setText(formatDoubleAsPercentage(billableHours,totalHours));
		
		// customer facing
		shape.getCell(2, 1).setText(formatDoubleAsPercentage(monthCustomerFacing[0],monthTotal[0]));
		shape.getCell(2, 2).setText(formatDoubleAsPercentage(monthCustomerFacing[1],monthTotal[1]));
		shape.getCell(2, 3).setText(formatDoubleAsPercentage(monthCustomerFacing[2],monthTotal[2]));
		shape.getCell(2, 4).setText(formatDoubleAsPercentage(customerFacingHours,totalHours));

		
//		for (int row = 1; row < numRows; row++) {
//			for (int col = 1; col < numCols; col++) {
//				shape.getCell(row, col).setText("U");
//
//			}
//		}
		
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

		// adjust rows to cover header row  + total row
		if (numRows < dept.length + 2) {
			for (int i = numRows; i < dept.length + 2; i++) {
				shape.addRow();
				for (int j = 0; j < numCols; j++) {
					shape.getRows().get(shape.getNumberOfRows() - 1).addCell();
				}
			}
		}
		numRows = shape.getNumberOfRows();
		String fmtString = "#,##0.000";
		
		XSLFTextParagraph p = shape.getCell(0,1).getTextParagraphs().get(0);
		p.getTextRuns().get(0).setText(selectedMonth + " YTD");

		double ytd = 0.0;
		double backlog = 0.0;
		double forecast = 0.0;
		for (int row = 1; row < numRows - 1 ; row++) {
			ytd = ytd + deptYTD[row - 1];
			backlog = backlog + deptBacklog[row - 1];
			forecast = forecast + deptForecast[row - 1];
			shape.getCell(row, 0).setText(deptNames.get(dept[row - 1]));
			shape.getCell(row, 1).setText(formatDouble(deptYTD[row - 1]/1000000, fmtString));
			shape.getCell(row, 2).setText(formatDouble(deptBacklog[row - 1]/1000000, fmtString));
			shape.getCell(row, 3).setText(formatDouble(deptForecast[row - 1]/1000000, fmtString));
		}
		shape.getCell(numRows - 1, 0).setText("Total Garage");
		shape.getCell(numRows - 1, 1).setText(formatDouble(ytd/1000000, fmtString));
		shape.getCell(numRows - 1, 2).setText(formatDouble(backlog/1000000, fmtString));
		shape.getCell(numRows - 1, 3).setText(formatDouble(forecast/1000000, fmtString));
		
	}
	
	private String formatDouble(double d, String formatStr) {
		return dataFormatter.formatRawCellContents(d, -1, formatStr);
	}
	
	private String formatDoubleAsPercentage(double n, double d) {
		String formatStr = "0%";
		if (d == 0.00) {
			return "n/a";
		} 
		return dataFormatter.formatRawCellContents(n/d, -1, formatStr);
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
			qtr[0] = m < 2 ? m : 2;
			break;

		case 1:
			qtr[0] = 2;
			qtr[1] = m < 5 ? m : 5;
			
			break;

		case 2:
			qtr[0] = 2;
			qtr[1] = 5;
			qtr[2] = m < 8 ? m : 8;
			
			break;

		case 3:
			qtr[0] = 2;
			qtr[1] = 5;
			qtr[2] = 8;
			qtr[3] = m < 11 ? m : 11;
			
			break;

		default:
			break;
		}
		
		for (int i = 0; i < 12; i++) {
			quarterlyTable[i] = 0.0;
		}
		
		for (int i = 0; i <= m; i++) {
			int qs = (int) (i/3);
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
			deptYTD[i] = 0.0;
			for (int j = 0; j < deptMonthly[i].length; j++) {
				deptMonthly[i][j] = 0.0;
				double val = 0.0;
				for (int k = 0; k < 3; k++) {
					CellReference cref = new CellReference("Garage & GEO Summary", pRow + k, pCol + j, true, true);
					val = getNumericCellValue(wb, cref.formatAsString());
					deptMonthly[i][j] = deptMonthly[i][j] + val;
					
					if (j <= getMonthNumber(selectedMonth)) {
						deptYTD[i] = deptYTD[i] + val; 
					}
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


	public String getPptxTemplate() {
		return pptxTemplate;
	}


	public void setPptxTemplate(String pptxTemplate) {
		this.pptxTemplate = pptxTemplate;
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


	private String getPptxOutput() {
		return pptxOutput;
	}


	private void setPptxOutput(String pptxOutput) {
		this.pptxOutput = pptxOutput;
	}


	private String getUtilizationXlsxPath() {
		return utilizationXlsxPath;
	}


	private void setUtilizationXlsxPath(String utilizationXlsxPath) {
		this.utilizationXlsxPath = utilizationXlsxPath;
	}


	public String getSelectedYear() {
		return selectedYear;
	}


	public void setSelectedYear(String selectedYear) {
		this.selectedYear = selectedYear;
	}
	
	
}
