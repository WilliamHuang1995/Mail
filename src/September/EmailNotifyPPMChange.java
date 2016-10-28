package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

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
import com.agile.api.ProgramConstants;
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
	private static HashMap<String, User> userList = new HashMap<String, User>();
	private static Properties props;

	public EmailNotifyPPMChange() {
		ini = new Ini();
		log = new LogIt("PPM");
		try {
			log.setLogFile(ini.getValue("File Location", "LOG_FILE_PATH")+"log");
		} catch (IOException e1) {
			e1.printStackTrace();
			log.log("LOG_FILE_PATH有誤，請修改之後再重新嘗試");
		}
		log.log("讀取Config里的[Server Info]...");
		username = ini.getValue("Server Info", "username");
		password = ini.getValue("Server Info", "password");
		connectString = ini.getValue("Server Info", "URL");
		log.log(1,"成功!");
		// Read Excel File
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		filename = format.format(today) + ".xlsx";
		fileLocation = ini.getValue("File Location", "FILE_PATH");
		try {
			fis = new FileInputStream(new File(fileLocation + filename));
		} catch (Exception e) {
			log.log("找不到Excel檔案，請檢查File Location里面的FILE_PATH");
			System.exit(1);
		}
		
		// initialize email properties
		props = initializeProperties();
	}

	public static void main(String[] args) {
		//初始化
		EmailNotifyPPMChange ini = new EmailNotifyPPMChange();
		
		try {
			log.log("嘗試登入Agile PLM環境。。");
			session = WebClient.getAgileSession(username, password, connectString);
			if (session != null)
				log.log(1,"登入成功");

		} catch (APIException e) {
			log.log("登入失敗，請確認[Server Info]的資料");
			System.exit(1);
		}

		try {
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();
			// skip first row
			rowIterator.next();

			// read sheet
			log.log("讀取Excel裡面的資料。。。");
			while (rowIterator.hasNext()) {
				HSSFRow row = (HSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				String role = cellIterator.next().getStringCellValue();// role
				String name = cellIterator.next().getStringCellValue();// name
				String userID = cellIterator.next().getStringCellValue();// userid
				String email = cellIterator.next().getStringCellValue();// email
				String programName = cellIterator.next().getStringCellValue();// project name
				String URL = cellIterator.next().getStringCellValue();// url
				String assignedBy = cellIterator.next().getStringCellValue();//assigned by
				log.log("角色: "+role+" 名字: "+name+" ID: "+userID+" 信箱: "+email+" 方案ID: "+programName+" 指派人: "+assignedBy);
				// Check if user is the current updated.
				IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, programName);
				ICell cell = program.getCell("Page Three."+role);
				String userInCell = cell.getValue().toString();
				log.log("檢查使用者是否是當下指派的員工。。	");
				
				// if user currently is in the role
				if (StringUtils.contains(userInCell, userID)) {
					log.log(1,name+" 是當下指派的員工");
					if (userList.get(name) != null) {
						User user = userList.get(name);
						user.addNewProject(programName, URL, assignedBy);
						//adds new role and also check if it is already added
						user.addNewRole(programName, role);
					} else {
						User user = new User(name, userID, email, programName, URL, assignedBy, role);
						userList.put(name, user);
					}
				}else{
					log.log(1,name+" 不是當下指派的員工，這代表在發這封郵件之前他/她曾經有被指派到這個角色但是又被移除了");
					log.log(1,"跳過。");
				}
			}
			fis.close();
			if(userList.size()==0){
				log.log("所有的員工都並非當下指派的人，系統即將結束");
				System.exit(0);
			}
			log.log("已讀完Excel檔，準備整理訊息并發送郵件給相關人士");
			
			sendMail(props);

		} catch (APIException | IOException e) {
			e.printStackTrace();
			System.exit(1);
		}

	}

	private Properties initializeProperties() {
		Properties props = new Properties();
		try {
			log.log("初始化郵件屬性...");
			String transportProtocol = ini.getValue("Admin Mail", "transport protocol");
			String host = ini.getValue("Admin Mail", "host");
			String protocolPort = ini.getValue("Admin Mail", "protocol port");
			log.log(1, "Transport Protocol: " + transportProtocol);
			log.log(1, "Mail host         : " + host);
			log.log(1, "Protocol Port     : " + protocolPort);
			props.setProperty("mail.transport.protocol", transportProtocol);
			props.setProperty("mail.host", host);
			props.setProperty("mail.protocol.port", protocolPort);
			props.setProperty("mail.smtp.auth", "true");

		} catch (Exception e) {
			e.printStackTrace();
			log.log("系統在設定寄件性能時出錯,請確認[Admin Mail]裡的設定都是正確再重試");
			System.exit(1);
		}
		return props;
	}

	private static void testRunMail(String username, String password, Properties props) {
		Session mailSession = Session.getDefaultInstance(props, null);
		log.log("測試是否能連接到MAIL SERVER");
		try {
			Transport transport = mailSession.getTransport();
			transport.connect(username, password);
			transport.close();

			log.log("成功連接MAIL");

		} catch (NoSuchProviderException e) {
			log.log("登入失敗,請查看 [Admin Mail] 的前3項是否正確");
			System.exit(1);
		} catch (MessagingException e) {
			log.log("登入失敗,請查看 [Admin Mail] 的設定是否正確");
			System.exit(1);
		}

	}

	private static void sendMail(Properties props) {
		try {
			// determine if config wants logo
			String logo = ini.getValue("Settings", "logo");
			boolean useLogo = logo.equalsIgnoreCase("yes");

			String username = ini.getValue("Admin Mail", "username");
			String password = ini.getValue("Admin Mail", "password");
			testRunMail(username, password, props);
			Iterator iter = userList.entrySet().iterator();
			log.log("開始發送郵件通知相關人員");
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();
				String key = (String) pair.getKey();
				String userEmail = "william@anselm.com.tw";

				Session mailSession = Session.getDefaultInstance(props, null);
				Transport transport = mailSession.getTransport();
				// 產生整封 email 的主體 message
				MimeMessage message = new MimeMessage(mailSession);

				// 設定主旨
				String subject = "New Changes in your role in Project";
				message.setSubject(subject, "utf-8");
				MimeBodyPart textPart = new MimeBodyPart();
				StringBuffer html = new StringBuffer();

				html.append("<!DOCTYPE html><html><head><style>" + "table,th,td{border: 1px solid black; }"
						+ "td,th{text-align:center;}" + "</style></head><body>");
				html.append("<p>" + pair.getKey() + ", 您的PLM角色有變更</p>");
				if (useLogo) {
					html.append("<img src='cid:image'/><br>");
				}

				html.append("<table><tr><td>Project Name</td><td>Project ID</td><td>Assigned By</td><td>Role(s) Assigned</td><td>Link</td></tr>");
				html.append(pair.getValue());
				html.append("</table></body></html>");
				html.append("<p>Sincerely,</p><p></p><p>Your Agile PLM Administrator</p>");
				textPart.setContent(html.toString(), "text/html; charset=UTF-8");

				Multipart email = new MimeMultipart();
				email.addBodyPart(textPart);
				// Oracle Logo
				if (useLogo) {
					MimeBodyPart picturePart = new MimeBodyPart();
					FileDataSource fds = new FileDataSource("logo.png");
					picturePart.setDataHandler(new DataHandler(fds));
					picturePart.setFileName(fds.getName());
					picturePart.setHeader("Content-ID", "<image>");
					email.addBodyPart(picturePart);
				}
				message.setContent(email);
				// replace wjhuang@ucsd.edu with userEmail
				message.addRecipient(Message.RecipientType.TO, new InternetAddress("william@anselm.com.tw"));
				message.setFrom(new InternetAddress(username)); // 寄件者
				transport.connect(username, password);
				transport.sendMessage(message, message.getRecipients(Message.RecipientType.TO));
				log.log("郵件成功發送給: " + userEmail);
				transport.close();

			}
		} catch (AddressException e) {
			e.printStackTrace();
		} catch (NoSuchProviderException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		}

	}

	public static class User {
		private String name;
		private String userID;
		private String email;
		private HashMap<String, String[]> projects = new HashMap<String, String[]>(); // project and link
		private HashMap<String, String> roles = new HashMap<String,String>();// project and role

		public User() {

		}

		public User(String name, String userID, String email, String project, String URL, String assignedBy, String role) {
			this.setName(name);
			this.setEmail(email);
			this.setUserID(userID);
			projects.put(project, new String[]{URL,assignedBy});
			roles.put(project,role);
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

		public void addNewProject(String project, String URL, String assignedBy) {
			projects.put(project, new String[]{URL,assignedBy});
		}
		
		public void addNewRole(String project, String role) {
			if (roles.get(project)==null){
				roles.put(project, role);
				return;
			}
			String newRole = roles.get(project);
			if(!StringUtils.contains(newRole, role))
				newRole = newRole+" "+ role;
			roles.put(project, newRole);
		}

		@Override
		public String toString() {
			String toReturn = "";
			Iterator iter = projects.entrySet().iterator();
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();

				String projectName = "did not find project";
				try {
					IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, pair.getKey());
					projectName = (String) program.getValue(ProgramConstants.ATT_GENERAL_INFO_NAME);
				} catch (APIException e) {
					e.printStackTrace();
				}
				toReturn = toReturn + "<tr><td>" + projectName + "</td><td>" + pair.getKey() +"</td><td>"+((String[])pair.getValue())[1]+"</td><td>"+roles.get(pair.getKey())+"</td><td><a href="
						+ ((String[])pair.getValue())[0] + ">Link to project</a></td></tr>";
			}

			return toReturn;
		}

	}
}
