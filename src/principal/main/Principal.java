package principal.main;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Principal {

	public static void main(String[] args) throws IOException {


		FileInputStream arquivo = new FileInputStream(new File(
				"C:\\Users\\Madson\\Documents\\Pasta1.xlsx"));
		
		

		XSSFWorkbook workbook = new XSSFWorkbook(arquivo);
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		try {
		    workbook.write(bos);
		} finally {
		    bos.close();
		}
		byte[] bytes = bos.toByteArray();
		
		InputStream stream = new ByteArrayInputStream(bytes);
		XSSFWorkbook workbook2 = new XSSFWorkbook(stream);
		
		
		XSSFSheet sheet1 = workbook2.getSheetAt(0);
		
		Iterator<Row> rowIterator = sheet1.iterator();
		
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();//linha
			Cell cell  = row.getCell(0);
			Cell cell1  = row.getCell(1);
			System.out.println(cell.getStringCellValue());
			System.out.println(cell1.getNumericCellValue());
		}
		workbook.close();
		
	}

}
