package marathishaadi.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
/**
 * it's used to read data from Excel-workBok, based on user argument..
 * 
 * @param SheetName
 * @param rowNum
 * @param colNum
 * @return data
 */
public class ExcelData {
	
	XSSFWorkbook workbook;
    XSSFSheet Data;
    
public ExcelData(String excellpath)
{
	try {
		File src=new File("./data/marathishadi.xlsx");   
		// Load the file.
		FileInputStream fis = new FileInputStream(src);
		// Load he workbook.
		workbook = new XSSFWorkbook(fis);
		// Load the sheet in which data is stored.
		Data=workbook.getSheetAt(0);
			}catch (Exception e)
			{
				System.out.println(e.getMessage());
			}}

public String getData(String SheetName,int rowNum,int colNum)
{
	String data=Data.getRow(rowNum).getCell(colNum).getStringCellValue();
	return data;
	
	}
/**
 * its used the property key value from marathi.properties
 * @param key
 * return value
 * throws Throwable
 */
public String getPropetyKeyvalue(String key) throws IOException
{
	FileInputStream fis=new FileInputStream("./data/marathi.properties");
	Properties pobj=new Properties();
	pobj.load(fis);
	String value=pobj.getProperty(key);
	return value;
	}
}
