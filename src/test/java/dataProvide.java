import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

public class dataProvide {

	DataFormatter formatter = new DataFormatter();

	@Test(dataProvider = "driveTest")
	public void testcasedata(String greeting, String communication, String id) {

		System.out.println(greeting + communication + id);
	}

	@DataProvider(name = "driveTest")
	public Object[][] getData() throws IOException {
//		Object[][] data = {{"Hello","text",1},{"bye","message",2},{"solo","call",3}};
//		return data;

		FileInputStream fis = new FileInputStream("D://dataDrivenExcel.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(0);
		int rowsCount = sheet.getPhysicalNumberOfRows();
		XSSFRow row = sheet.getRow(0);
		int colCount = row.getLastCellNum();
		Object data[][] = new Object[rowsCount - 1][colCount];
		for (int i = 0; i < rowsCount - 1; i++) {
			row = sheet.getRow(i + 1);
			for (int j = 0; j < colCount; j++) {
				XSSFCell cell = row.getCell(j);
				data[i][j] = formatter.formatCellValue(cell);
			}
		}
		return data;
	}

}
