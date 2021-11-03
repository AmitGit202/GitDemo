package dataDriven.ExcelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelTest
{
	@Test
	public void getData() throws IOException
	{
		FileInputStream fis = new FileInputStream("D:\\Software\\eclipse-jee-2019-06-R-win32-x86_64\\eclipse\\WorkSpace\\ExcelDataProvider\\ExcelDataDriven\\DataSheet.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet =wb.getSheetAt(0);
		int rowCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		Object data[][]= new Object[rowCount-1][colCount];

		for(int i=0;i<rowCount-1;i++)
		{
			row = sheet.getRow(i+1);
			for(int j=0; j<colCount;j++)
			{
				System.out.println(row.getCell(j));
			}
			System.out.println("Outer loop Ended here");
		}



	}

}
