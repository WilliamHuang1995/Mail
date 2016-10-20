package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.anselm.plm.utilobj.Ini;

public class WriteSheet {

	public static void main(String[] args) throws Exception {
		System.out.println(StringUtils.difference("[user1,user2,user3]", "[user1,user2,user3,user4]"));

		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		String filename = format.format(today) + ".xlsx";
		Ini ini = new Ini();
		String fileLoc = ini.getValue("File Location", "FILEPATH");

		File temp = new File(fileLoc + filename);
		boolean exists = temp.exists();
		System.out.println(exists);
		HSSFSheet spreadsheet = null;
		HSSFWorkbook workbook = null;
		int rowcount = 0;
		if (exists) {
			FileInputStream fis = new FileInputStream(new File(fileLoc + filename));
			workbook = new HSSFWorkbook(fis);
			spreadsheet = workbook.getSheetAt(0);
			// used to append to the current file
			rowcount = spreadsheet.getLastRowNum();

		} else {
			// Create blank workbook
			workbook = new HSSFWorkbook();
			// Create a blank sheet
			spreadsheet = workbook.createSheet(" Employee Info ");
			// Create row object
		}

		// This data needs to be written (Object[])
		Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
		if(!exists){
			empinfo.put("1", new Object[] { "User ID", "Email", "Name", "Project Name" });
		}
		empinfo.put("2", new Object[] { "tp01", "a@a.com", "Gopal", "Technical Manager" });
		empinfo.put("3", new Object[] { "tp02", "b@b.com", "Manisha", "Proof Reader" });
		empinfo.put("4", new Object[] { "tp03", "c@c.com", "Masthan", "Technical Writer" });
		empinfo.put("5", new Object[] { "tp04", "d@d.com", "Satish", "Technical Writer" });
		empinfo.put("6", new Object[] { "tp05", "e@e.com", "Krishna", "Technical Writer" });
		// Iterate over data and write to sheet
		Set<String> keyid = empinfo.keySet();
		for (String key : keyid) {
			HSSFRow row = spreadsheet.createRow(rowcount++);
			Object[] objectArr = empinfo.get(key);
			int cellid = 0;
			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}
		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File(fileLoc + filename));
		workbook.write(out);
		workbook.close();
		out.close();
		System.out.println("Write Successful");

	}
}