import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class UpdateExcelSheet {

	public static void main(String[] args) {
		FileInputStream datafile = null;
		FileOutputStream fos =null;
		Workbook workbook;
		Sheet sheet;
		
		System.out.println("Enter Row No to update(separate with comma):-");
		
		try 
		{
			datafile = new FileInputStream("src/test/resources/selenium/TestData.xls");
			workbook = new HSSFWorkbook(datafile);
			sheet=workbook.getSheet("TestConfig");
			fos=new FileOutputStream("src/test/resources/selenium/TestData.xls");
			for (String string : args) 
			{
				int rowNo=Integer.parseInt(string);
				String cellValue=sheet.getRow(rowNo).getCell(2).getStringCellValue();
				System.out.println("Current Value:-"+cellValue);
				if(cellValue.equals("N"))
				{
					sheet.getRow(rowNo).createCell(2).setCellValue("Y");
				}
				else if(cellValue.equals("Y"))
				{
					sheet.getRow(rowNo).createCell(2).setCellValue("N");
				}
			}
			
			workbook.write(fos);
			
		} 
		catch (Exception e) 
		{
			// TODO: handle exception
		}
		finally
		{
			try 
			{
				datafile.close();
				fos.close();
			} 
			catch (IOException e) 
			{
				e.printStackTrace();
			}
		}

	}

}
