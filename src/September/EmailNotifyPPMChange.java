package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IProgram;
import com.agile.api.IProject;
import com.agile.api.ITable;
import com.agile.api.IUser;
import com.agile.api.TableTypeConstants;
import com.agile.api.UserConstants;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;

import util.WebClient;

public class EmailNotifyPPMChange {
	private static Ini ini = null;
	private static LogIt log = null;
	private static IAgileSession session;
	private String filename;
	private String fileLocation;
	private static FileInputStream fis = null;
	private static String username;
	private static String password;
	private static String connectString;
	private static HashMap<String,User> userList = new HashMap<String,User>();

	public EmailNotifyPPMChange() {
		ini = new Ini();
		log = new LogIt("PPM");
		username = ini.getValue("Server Info", "username");
		password = ini.getValue("Server Info", "password");
		connectString = ini.getValue("Server Info", "URL");

		// Read Excel File
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		filename = format.format(today) + ".xlsx";
		fileLocation = ini.getValue("File Location", "FILEPATH");
		try {
			fis = new FileInputStream(new File(fileLocation + filename));
		} catch (Exception e) {
			log.log("can't find file");
			System.exit(1);
		}
	}

	public static void main(String[] args) {
		//Initialize
		EmailNotifyPPMChange ini = new EmailNotifyPPMChange();
		try {
			session = WebClient.getAgileSession(username, password, connectString);
			if (session != null)
				log.log("Logged in");
			
		} catch (APIException e) {
			log.log("Failure to log in");
			System.exit(1);
		}
		
		try {
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();
			//skip first row
			rowIterator.next();
			
			//read sheet
			while (rowIterator.hasNext()) {
				HSSFRow row = (HSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				String role = "Page Three."+cellIterator.next().getStringCellValue();//role
				String name = cellIterator.next().getStringCellValue();//name
				String userID = cellIterator.next().getStringCellValue();//userid
				String email = cellIterator.next().getStringCellValue();//email
				String programName = cellIterator.next().getStringCellValue();//project name
				String URL = cellIterator.next().getStringCellValue();//url
				
				//Check if user is the current updated.
				IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, programName);
				ICell cell = program.getCell(role);
				String userInCell = cell.getValue().toString();
				
				//if user currently is in the role
				if(StringUtils.contains(userInCell, userID)){
					log.log(name+" is currently the assigned user in role: "+role +" for Project: "+programName);
					if(userList.get(name)!= null){
						User user = userList.get(name);
						user.addNewProject(programName, URL);
					}
					else{
						User user = new User(name,userID,email,programName,URL);
						userList.put(name,user);
					}
				}
			}
			
			fis.close();
			
			


		} catch (APIException | IOException e) {
			e.printStackTrace();
			System.exit(1);
		}

	}
	
	public static class User{
		private String name;
		private String userID;
		private String email;
		private HashMap<String,String> projects = new HashMap<String,String>(); //project and link
		
		public User(){
			
		}
		public User(String name, String userID, String email, String project, String URL){
			this.setName(name);
			this.setEmail(email);
			this.setUserID(userID);
			projects.put(project, URL);
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		public String getUserID() {
			return userID;
		}
		public void setUserID(String userID) {
			this.userID = userID;
		}
		public String getEmail() {
			return email;
		}
		public void setEmail(String email) {
			this.email = email;
		}
		public void addNewProject(String project, String URL){
			projects.put(project, URL);
		}
		@Override
		public String toString(){
			return name+" "+userID;
		}
		
	}
}
