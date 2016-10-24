package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
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
		String str = "Administrator, admin (admin);Gabriele Braga, Arrow Electronics (1ae001.01);Julie Fozard, Channel One Limited (1co001.01)";
		String str2 = "Administrator, admin (admin);Gabriele Braga, Arrow Electronics (1ae001.01);Julie Fozard, Channel One Limited (1co001.01);Pang, Philip, Avnet Electronics Marketing Australia (1ae002.01)";
		ArrayList<String> oldValue = new ArrayList(Arrays.asList(str.split("\\s*;\\s*")));
		ArrayList<String> newValue = new ArrayList<String>(Arrays.asList(str2.split("\\s*;\\s*")));
		newValue.removeAll(oldValue);
		System.out.println(oldValue);
		System.out.println(newValue);
		Iterator<String> it = newValue.iterator();
		while(it.hasNext()){
			String user= it.next();
			String userID = user.substring(user.lastIndexOf("(")+1,user.lastIndexOf(")"));
			System.out.println(userID);
		}
		

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
			empinfo.put("1", new Object[] { "User Role", "Email","Project Name" });
		}
		empinfo.put("2", new Object[] { "ME", "a@a.com", "Technical Manager" });
		empinfo.put("3", new Object[] { "EE", "b@b.com", "Proof Reader" });
		empinfo.put("4", new Object[] { "PM", "c@c.com", "Technical Writer" });
		empinfo.put("5", new Object[] { "BIOS", "d@d.com", "Technical Writer" });
		empinfo.put("6", new Object[] { "DQA", "e@e.com", "Technical Writer" });
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