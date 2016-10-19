package September;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;



public class WriteSheet {
	public static void main(String[] args) throws Exception {
		System.out.println(StringUtils.difference("[user1,user2,user3]", "[user1,user2,user3,user4]"));
		
		
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		String filename = format.format(today)+".xlsx";
		System.out.println("Current Date: "+format.format(today));
		// Create blank workbook
		HSSFWorkbook workbook = new HSSFWorkbook();
		// Create a blank sheet
		HSSFSheet spreadsheet = workbook.createSheet(" Employee Info ");
		// Create row object
		HSSFRow row;
		// This data needs to be written (Object[])
		Map<String, Object[]> empinfo = new TreeMap<String, Object[]>();
		empinfo.put("1", new Object[] { "User ID", "Name", "Project Name" });
		empinfo.put("2", new Object[] { "tp01", "Gopal", "Technical Manager" });
		empinfo.put("3", new Object[] { "tp02", "Manisha", "Proof Reader" });
		empinfo.put("4", new Object[] { "tp03", "Masthan", "Technical Writer" });
		empinfo.put("5", new Object[] { "tp04", "Satish", "Technical Writer" });
		empinfo.put("6", new Object[] { "tp05", "Krishna", "Technical Writer" });
		// Iterate over data and write to sheet
		Set<String> keyid = empinfo.keySet();
		int rowid = 0;
		for (String key : keyid) {
			row = spreadsheet.createRow(rowid++);
			Object[] objectArr = empinfo.get(key);
			int cellid = 0;
			for (Object obj : objectArr) {
				Cell cell = row.createCell(cellid++);
				cell.setCellValue((String) obj);
			}
		}
		// Write the workbook in file system
		FileOutputStream out = new FileOutputStream(new File(filename));
		workbook.write(out);
		out.close();
		System.out.println("Writesheet.xlsx written successfully");
	}
}