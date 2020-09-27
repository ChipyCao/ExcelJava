import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class poi {
	public static void main(String[] args) {
		try {
			int maxm = 17;
			int maxmindex = 0;
			
			int minm = 61;
			int minmindex = 0;
			
			int maxf = 17;
			int maxfindex = 0;
			
			int minf = 61;
			int minfindex = 0;
			
			//Creating Workbook in .xlsx format
			Workbook workbook = new XSSFWorkbook();

			// a) Creating the sheet
			Sheet sheet = workbook.createSheet("Ex.1");
		
			//Creating a top row with column headings
			String[] columnHeadings = {"Name","Age","Sex"};
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short)12);
			headerFont.setColor(IndexedColors.BLACK.index);
		
			//Creating a header CellStyle 
			CellStyle headerStyle = workbook.createCellStyle();
			headerStyle.setFont(headerFont);
			headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.index);
		
			//Creating header row
			Row headerRow = sheet.createRow(0);
		
			for(int i=0; i<columnHeadings.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(columnHeadings[i]);
				cell.setCellStyle(headerStyle);
			}
			
			//Filling data
			ArrayList<invoices> a = createData();
			int rownum =1;
			for(invoices i : a) {
				Row row = sheet.createRow(rownum++);
				row.createCell(0).setCellValue(i.getName());
				row.createCell(1).setCellValue(i.getAge());
				row.createCell(2).setCellValue(i.getSex());
				
				switch (i.getSex()) {
				  case "M":
				  		{
						if(i.getAge() > maxm)
									{
									maxm=i.getAge();
							 		maxmindex=rownum;}
								
						else if(i.getAge() < minm)
								{
								minm=i.getAge();
								minmindex=rownum;
								}
					    break;
						}

				  case "F":
				  		{
						if(i.getAge() > maxf)
							{
							maxf=i.getAge();
							maxfindex=rownum;
							}
						else if(i.getAge() < minf)
							{
							minf=i.getAge();
							minfindex=rownum;
							}
					    break;
						}
					}

				}
						
			
			////////////////////////////////////////////////         a) Highlighting the max & min           //////////////////////////////////////////////////////////////////////	
			
			CellStyle highlight = workbook.createCellStyle();
			highlight.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			highlight.setFillForegroundColor(IndexedColors.BLUE.index);
			
			for(int i=0; i<columnHeadings.length; i++) {			
				sheet.getRow(maxmindex).getCell(i).setCellStyle(highlight);
				sheet.getRow(minmindex).getCell(i).setCellStyle(highlight);
				sheet.getRow(maxfindex).getCell(i).setCellStyle(highlight);
				sheet.getRow(minfindex).getCell(i).setCellStyle(highlight);
			}
		
			
			////////////////////////////////////////////////        b)Creating the other two sheets			//////////////////////////////////////////////////////////////////////
			
			
			/////////      MALE     ////////
			Sheet Male = workbook.createSheet("Male");
			String[] columnHeadingM = {"Name","Age"};
			Row headerRowM = Male.createRow(0);
			
			for(int i=0; i<columnHeadingM.length; i++) {
				Cell cell = headerRowM.createCell(i);
				cell.setCellValue(columnHeadingM[i]);
				cell.setCellStyle(headerStyle);
			}
			int rownumM =1;
			int sm=0;
			for(invoices i : a) 
				if((i.getSex().equals("M")))
				{
				Row row = Male.createRow(rownumM++);
				row.createCell(0).setCellValue(i.getName());
				row.createCell(1).setCellValue(i.getAge());
				sm+=i.getAge();
				
				}

			Row avgM = Male.createRow(rownumM);
			Cell avgMC = avgM.createCell(1);
			avgMC.setCellFormula("AVERAGE(B2:B7)");
			avgMC.setCellStyle(headerStyle);
			
			for(int j=0;j<columnHeadingM.length-1;j++) {
				Male.autoSizeColumn(j);
			}
			
			
			/////////     FEMALE     ////////
			
			
			Sheet Female = workbook.createSheet("Female");
			String[] columnHeadingF = {"Name","Age"};
			Row headerRowF = Female.createRow(0);
			
			for(int i=0; i<columnHeadingF.length; i++) {
				Cell cell = headerRowF.createCell(i);
				cell.setCellValue(columnHeadingF[i]);
				cell.setCellStyle(headerStyle);
			}
			int rownumF =1;
			int sf=0;
			for(invoices i : a) 
				if((i.getSex().equals("F")))
				{
				Row row = Female.createRow(rownumF++);
				row.createCell(0).setCellValue(i.getName());
				row.createCell(1).setCellValue(i.getAge());
				sf+=i.getAge();
				}
			
			Row avgF = Female.createRow(rownumF);
			Cell avgFC = avgF.createCell(1);
			avgFC.setCellFormula("AVERAGE(B2:B4)");
			avgFC.setCellStyle(headerStyle);
			
			for(int j=0;j<columnHeadingF.length-1;j++) {
				Female.autoSizeColumn(j);
			}
			
			
			////////////////////////////////////////////////			Writing the output to file			//////////////////////////////////////////////////////////////////////
			
			FileOutputStream fileOut = new FileOutputStream("lib/Invoices.xlsx");
			workbook.write(fileOut);
			fileOut.close();			
			workbook.close();
			System.out.println("Completed");			
		}
		catch(Exception e) {
			e.printStackTrace();
		}
	}
	
	private static ArrayList<invoices> createData() {
	ArrayList<invoices> a = new ArrayList();
	a.add(new invoices("Calvin",29,"M"));
	a.add(new invoices("Luther",30,"M"));
	a.add(new invoices("Ioana",33,"F"));
	a.add(new invoices("Maria",21,"F"));
	a.add(new invoices("Andrei",60,"M"));
	a.add(new invoices("Natasha",19,"F"));
	a.add(new invoices("Alfred",53,"M"));
	a.add(new invoices("Faust",60,"M"));
	a.add(new invoices("Henry",45,"M"));
	a.add(new invoices("Jeanne",25,"F"));
	return a;
	}
	
}
