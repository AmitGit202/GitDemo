package dataDriven.ExcelDataProvider;

import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;


public class DataProvide 
{

	//Multiple set of data to our Tests
	//array public
	//5 sets of data as 5 arrays from data provider to your tests
	//Then your test will run 5 times with 5 separate sets of data(arrays)

	DataFormatter formatter = new DataFormatter();

	@Test(dataProvider="driventest")
	public void testCaseData(String greeting, String communication,String id)
	{
		System.out.println(greeting+communication+id);
	}
	/*	
	@DataProvider(name="driventest")
	public Object[][] getData()
	{
		Object[][] data = {{"hello","text",1},{"bye","message",2},{"solo","call",3}};
		return data;

	}
	 */
	@DataProvider(name="driventest")
	public Object[][] getData() throws IOException
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
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}

		}

		return data;

	}
}
