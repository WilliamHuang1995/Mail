package September;

import java.io.File;
import java.io.FileInputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.anselm.plm.utilobj.Ini;

public class ReadSheet {
	static HSSFRow row;

	public static void main(String[] args) throws Exception {
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		String filename = format.format(today)+".xlsx";
		Ini ini = new Ini();
		String fileLoc = ini.getValue("File Location", "FILEPATH");
		FileInputStream fis = null;
		try{
		fis = new FileInputStream(new File(fileLoc+filename));
		}catch(Exception e){System.out.println("can't find file");System.exit(1);}
		HSSFWorkbook workbook = new HSSFWorkbook(fis);
		HSSFSheet spreadsheet = workbook.getSheetAt(0);
		Iterator<Row> rowIterator = spreadsheet.iterator();
		//skip first row
		rowIterator.next();
		while (rowIterator.hasNext()) {
			row = (HSSFRow) rowIterator.next();
			Iterator<Cell> cellIterator = row.cellIterator();
			String name = cellIterator.next().getStringCellValue();//name
			String email = cellIterator.next().getStringCellValue();//email
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				/*switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue() + " \t\t ");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print(cell.getStringCellValue() + " \t\t ");
					break;

				}*/
			}
			System.out.println();
			System.out.println("User: "+name+", Email: "+email);
		}
		fis.close();
	}
}