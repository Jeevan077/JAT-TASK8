package readexcelsheet;



import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteReadExcel {
	
     //Using this method to Write the data in Excel sheet
	
	public void WriteExcel(String Sheetname, int rowNum, int cellNum, String desc) {
		//To open the file from outside
		//Declaring the object
		FileInputStream fis;
		XSSFWorkbook wb;
		
		try {
			//Selecting the file to open
			fis =new FileInputStream("Utils//Student.xlsx");
			//Getting the spreadsheet
			wb=new XSSFWorkbook(fis);
			//Opening that particular sheet
			XSSFSheet s = wb.getSheet(Sheetname);
			//Selecting the row
			XSSFRow r = s.getRow(rowNum);
			//Creating the cell to write the data in excel
			XSSFCell c=r.createCell(cellNum);
			//Setting the cell value
			c.setCellValue(desc);
			//Output the cell value to the file so using FileOutputStream to write the data in excel
			FileOutputStream fos=new FileOutputStream("Utils//Student.xlsx");
			//Using write method to fos object
			wb.write(fos);
			
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		}

    // Using this method to read the data from excel
	
	public String getExcelData(String sheetName, int rowNum, int colNum) {
		//Returning the value through getExcelData
		String retVal = null;
		try {
			//To read the data from outside using FileInputStream
			FileInputStream fis = new FileInputStream("Utils//Student.xlsx");
			// Getting the whole spreadsheet
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			//Getting that particular sheet
			XSSFSheet s = wb.getSheet(sheetName);
			//Getting the particular row
			XSSFRow r = s.getRow(rowNum);
			//Getting the particular coloumn
			XSSFCell c = r.getCell(colNum);
			//Reading the data from excel so using return value
			retVal = WriteReadExcel.getCellValue(c);
			fis.close();
			wb.close();

		} catch (FileNotFoundException e) 
		{
			e.printStackTrace();
		} catch (IOException e)
		{
			e.printStackTrace();
		}
		return retVal;
	}

	public static String getCellValue(XSSFCell c)
	{
		switch(c.getCellType())
		{
		// Converting Numeric values to String
		case NUMERIC:
			return String.valueOf(c.getNumericCellValue());  
		// Converting Boolean values to String	
		case BOOLEAN:
			return String.valueOf(c.getBooleanCellValue());
		//Using getStringCellValue from Apache POR 	to print the data in String 
		case STRING:
			return c.getStringCellValue(); //Reads String cell value from excel
		default:
			return c.getStringCellValue();
			
		}
	}


	public static void main(String[] args) {
		WriteReadExcel x=new WriteReadExcel();
		
		//Reading the data in Excel sheet
		for (int i=0;i<5;i++){    // i for row
			for (int j=0;j<3;j++){ // j for coloumn
			// Calling the getexcelData method to select which Sheet, rows and coloumns
			String s=x.getExcelData("Sheet1", i, j); 
			System.out.println(s);
				
			}
		}
		//Calling Write excel method
		// Writing the data in Excel sheet
		x.WriteExcel("Sheet1", 0, 3, "Pass/Fail");
		x.WriteExcel("Sheet1", 1, 3, "Fail");
		x.WriteExcel("Sheet1", 2, 3, "Pass");
		x.WriteExcel("Sheet1", 3, 3, "Pass");	
		x.WriteExcel("Sheet1", 4, 3, "Pass");
	}
	
}

//Output:

/*
 * Name
Age
Email
John  Doe
30.0
john@test.com
Jane Doe
28.0
john@test.com
Bob Smith
35.0
jacky@example.com
Swapnil
37.0
swapnil@example.com
 */

