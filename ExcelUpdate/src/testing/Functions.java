package testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Functions {
	
	@SuppressWarnings("resource")
	public static int readIntData(int r,int c, String sheet, String file) throws IOException {
		int n;
		File f= new File(file);
		FileInputStream fis= new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis); 
		XSSFSheet sh= wb.getSheet(sheet);
		XSSFRow row= sh.getRow(r);
		XSSFCell cell= row.getCell(c);
		n= (int) cell.getNumericCellValue();
		return n;
	}
	
	@SuppressWarnings({ "resource" })
	public static String readData(int r,int c, String sheet, String file) throws IOException {
		String str = null;
		Date date= null;
		int st;
		File f= new File(file);
		FileInputStream fis= new FileInputStream(f);
		XSSFWorkbook wb = new XSSFWorkbook(fis); 
		XSSFSheet sh= wb.getSheet(sheet);
		XSSFRow row= sh.getRow(r);
		XSSFCell cell= row.getCell(c);
		CellType type = cell.getCellType();
		//cell.getCellStyle().getDataFormatString();
		
		 if(type == CellType.NUMERIC) {
			if (cell.getCellStyle().getDataFormatString().contains("%")) {
				float d= (float) cell.getNumericCellValue();
				int d1= (int) (d*100) ;
				str= d1+"%";  
			}
			else if(DateUtil.isCellDateFormatted(cell)) {
				date= cell.getDateCellValue();
				if(date!=null) {
					java.sql.Date date1 = new java.sql.Date(date.getTime());
					DateFormat df=new SimpleDateFormat("dd-MM-yyyy");
					str= df.format(date1);
				}
			}
			else {
				st= (int) cell.getNumericCellValue();
				str= String.valueOf(st);
			}
		}
		else if(type == CellType.STRING) {
			str= cell.getStringCellValue();
		}
		else
		{
			if(type == CellType.BLANK) {
			  str= null;
			}
			else {
				str= cell.getStringCellValue();
			}
		}
		return str;
	}
	
	@SuppressWarnings("resource")
	public static boolean setCellData(String filePathWithextension,String sheetName, int rowNum, int colNum, String data) {
        try {
               
            if (rowNum <= 0)
                      return false;
              File file = new File(filePathWithextension);
             XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(new FileInputStream(file)));
             XSSFSheet  sheet = workbook.getSheet(sheetName);
              XSSFRow row = sheet.getRow(1);
              CellStyle style = workbook.createCellStyle();  
              // Setting Background color  
              style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());  
              style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
              
               if (colNum == -1)
                     return false;
               //sheet.autoSizeColumn(colNum);
               row = sheet.getRow(rowNum - 1);
               if (row == null)
                     row = sheet.createRow(rowNum - 1);
              XSSFCell cell = row.getCell(colNum);
               if (cell == null)
                     cell = row.createCell(colNum);
               if(data!=null) {
            	   cell.setCellValue(data.trim().toString());
               }
               cell.setCellStyle(style);
               FileOutputStream fileOut = new FileOutputStream(file);
               workbook.write(fileOut);
               fileOut.close();
        } catch (Exception e) {
               e.printStackTrace();
               return false;
        }
        return true;
	}
	
	public static int mapColumn(String s) {
		 int result = 0; 
		    for (int i = 0; i < s.length(); i++) 
		    { 
		        result *= 26; 
		        result += s.charAt(i) - 'A' + 1; 
		    } 
		    return result-1; 
	}
	
	public static int findRow(String x, int start, int end, String sheet, String file) throws IOException {
		int targetRow=0;
		for(int i=start-1;i<=end-1;i++) {
			String y= readData(i, 0, sheet, file);
			if(y.equals(x)) {
				targetRow= i+1;	
			}
		}
		
		return targetRow;
	}
	
}