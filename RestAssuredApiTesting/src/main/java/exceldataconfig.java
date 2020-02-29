

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.hssf.util.HSSFColor.HSSFColorPredefined;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class exceldataconfig {
	
	
	
	XSSFWorkbook wb;
	XSSFSheet sheet1;
	public static FileInputStream fis2;
	public static XSSFWorkbook wb2;
	public static XSSFSheet ws;
	
	
	public exceldataconfig(String excelpath)
	{
        try {
			 File src = new File(excelpath);
			
			FileInputStream fis = new FileInputStream(src);
			
			
			
			wb = new XSSFWorkbook(fis); 
			
			
		}  
           catch (Exception e) {
        
        System.out.println(e.getMessage());	   
			
		}
	}
	
	
	public int getRowCount(String xlfile,String xlsheet) throws IOException
	{
		fis2=new FileInputStream(xlfile);
		wb2 = new XSSFWorkbook(fis2); 
		ws=wb2.getSheet(xlsheet);
		int rowcount = ws.getLastRowNum();
		wb2.close();
		fis2.close();
		return rowcount;
	}
	
	public int getCellCount() throws IOException
	{
		sheet1=wb.getSheet("Sheet1");
		XSSFRow row = sheet1.getRow(1);
		int cellcount = row.getLastCellNum();
		wb.close();
		return cellcount;
	}
	
	public String getData(int sheetnumber, int row, int col)
	{
		sheet1 =wb.getSheetAt(sheetnumber);
		String data = sheet1.getRow(row).getCell(col).getStringCellValue();
		return data; 
	}
	
	
	public double getnumericData(int sheetnumber, int row, int col)
	{
		sheet1 =wb.getSheetAt(sheetnumber);
		double data = sheet1.getRow(row).getCell(col).getNumericCellValue();
		return data; 
	}
	
	public String removeLastChar(String str) {
        return str.substring(0,str.length()-2);
    }
	
	
	public void writeData(int rnum, int cnum, String val, String clr) throws Exception
	
	{
		CellStyle color = wb.createCellStyle();
		if (clr.equals("red"))
		{
			//HSSFColor clrr = new HSSFColor();
        Font font = wb.createFont();
        font.setColor(HSSFColorPredefined.RED.getIndex());
        color.setFont(font);
		}
		if (clr.equals("green"))
		{
        
        Font font2 = wb.createFont();
        font2.setColor(HSSFColorPredefined.GREEN.getIndex());
        color.setFont(font2);
		}
		
		sheet1 =wb.getSheetAt(0);
        
		File src2 = new File("C:\\exceldata\\empdata.xlsx");
        sheet1.getRow(rnum).createCell(cnum).setCellValue(val);
        sheet1.getRow(rnum).getCell(cnum).setCellStyle(color);
        
		FileOutputStream fout = new FileOutputStream(src2); 
		
		wb.write(fout);
	}
	
	
	
	
public void writenumericData(int rnum, int cnum, long val, String clr) throws Exception
	
	{
		CellStyle color = wb.createCellStyle();
		if (clr.equals("red"))
		{
		
        Font font = wb.createFont();
        font.setColor(HSSFColorPredefined.RED.getIndex());
        color.setFont(font);
		}
		if (clr.equals("green"))
		{
        
        Font font2 = wb.createFont();
        font2.setColor(HSSFColorPredefined.GREEN.getIndex());
        color.setFont(font2);
		}
		
		sheet1 =wb.getSheetAt(0);
        
		File src2 = new File("C:\\exceldata\\empdata.xlsx");
        sheet1.getRow(rnum).createCell(cnum).setCellValue(val);
        sheet1.getRow(rnum).getCell(cnum).setCellStyle(color);
        
		FileOutputStream fout = new FileOutputStream(src2); 
		
		wb.write(fout);
	}
	
	
	
	
	
	
	
	
	

}
