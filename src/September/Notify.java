package September;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
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
import javax.mail.SendFailedException;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.swing.JOptionPane;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.agile.api.APIException;
import com.agile.api.IAgileList;
import com.agile.api.IAgileSession;
import com.agile.api.ICell;
import com.agile.api.IProgram;
import com.agile.api.IProject;
import com.agile.api.IUser;
import com.agile.api.ProgramConstants;
import com.agile.api.UserConstants;
import com.anselm.plm.util.AUtil;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;

import util.WebClient;

public class Notify {
	private static Ini ini = null;
	private static LogIt log = null;
	private static IAgileSession session;
	private static String filename;
	private static String fileLocation;
	private static FileInputStream fis = null;
	private static String username;
	private static String password;
	private static String connectString;
	private static HashMap<String, User> userList = new HashMap<String, User>();
	private static Properties props;
	static final String SUBADDRESS = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ChangeHandler&classid=6000&objid=";
	static final String PSR = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=PSRHandler&classid=4878&objid=";
	static final String QCR = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=QCRHandler&classid=4928&objid=";
	static final String PPM = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ActivityHandler&classid=18022&objid=";

	/*
	 * Inner Class
	 * 創造目的： 由於三個部分有可能衝突，所以建立此Class來追蹤一個user所有的流程
	 * @param name 用戶名
	 * @param userID 用戶的ID
	 * @param email 用戶的郵箱
	 * @param incompleteWorkflow 為簽核的流程
	 * @param overdueTask 為完成的工作
	 * @param projects 記錄指派人以及專案名
	 * @param roles 記錄專案被指派的角色
	 */
	public static class User {
		private String name;
		private String userID;
		private String email;
		private String incompleteWorkflow;
		private String overdueTask;
		public HashMap<String, String[]> projects = new HashMap<String, String[]>();
		public HashMap<String, String> roles = new HashMap<String, String>();
		

		public User() {
		}

		public User(String name, String userID, String email, String project, String URL, String assignedBy,
				String role) {
			this.setName(name);
			this.setEmail(email);
			this.setUserID(userID);
			projects.put(project, new String[] { URL, assignedBy });
			roles.put(project, role);
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

		public String getWorkflow() {
			return incompleteWorkflow;
		}

		public void setWorkflow(String workflow) {
			this.incompleteWorkflow = workflow;
		}

		public void addNewProject(String project, String URL, String assignedBy) {
			projects.put(project, new String[] { URL, assignedBy });
		}

		public void addNewRole(String project, String role) {
			if (roles.get(project) == null) {
				roles.put(project, role);
				return;
			}
			String newRole = roles.get(project);
			if (!StringUtils.contains(newRole, role))
				newRole = newRole + " " + role;
			roles.put(project, newRole);
		}

		public String getOverdueWork() {
			return overdueTask;
		}

		public void setOverdueWork(String overdueWork) {
			this.overdueTask = overdueWork;
		}

		public String toPPMString() {
			String toReturn = "<table><tr><td>Project Name</td><td>Project ID</td><td>Assigned By</td><td>Role(s) Assigned</td><td>Link</td></tr>";
			Iterator iter = projects.entrySet().iterator();
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();

				String projectName = "did not find project";
				try {
					IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, pair.getKey());
					projectName = (String) program.getValue(ProgramConstants.ATT_GENERAL_INFO_NAME);
				} catch (APIException e) {
					log.log(e);
					displayError();
				}
				toReturn += "<tr><td>" + projectName + "</td><td>" + pair.getKey() + "</td><td>"
						+ ((String[]) pair.getValue())[1] + "</td><td>" + roles.get(pair.getKey()) + "</td><td><a href="
						+ ((String[]) pair.getValue())[0] + ">Link to project</a></td></tr>";
			}
			toReturn += "</table>";
			return toReturn;
		}

		public String toPCString() {
			String toReturn = "<table><tr><td>表單類別</td><td>表單編號</td><td>表單描述</td><td>站別</td><td>已持續時間(天)</td></tr>";
			toReturn += this.getWorkflow() + "</table>";

			return toReturn;
		}

		public String toOverdueString() {
			String toReturn = "<table><tr><td>工作類別</td><td>工作名字</td><td>預計結束時間</td><td>已持續時間(天)</td></tr>";
			toReturn += this.getOverdueWork() + "</table>";
			return toReturn;
		}

	}
	/*
	 * 功能：初始化 將Log 與 Ini 初始化 初始化 Log 的路徑 （路徑取決於Config裏的LOG_FILE_PATH） LOG檔案名稱為
	 * 日期+Notify。log （EX: 1115Notify.log 代表11月15日所產生的log檔）
	 * 初始化客戶端資料，嘗試登入，若失敗將退出程式
	 */
	public Notify() {
		ini = new Ini();
		log = new LogIt("Notify");
		Date today = new Date();
		SimpleDateFormat format = new SimpleDateFormat("MMdd");
		String date = format.format(today);
		try {
			// 設定路徑
			log.setLogFile(ini.getValue("File Location", "LOG_FILE_PATH") + date + "Notify.log");
		} catch (IOException e) {
			log.log(e);
			log.log("LOG_FILE_PATH有誤，請修改之後再重新嘗試");
			displayError();
			System.exit(1);
		}

		// 讀取客戶端admin資料
		log.log("讀取Config里的[AgileAP]...");
		username = ini.getValue("AgileAP", "username");
		password = ini.getValue("AgileAP", "password");
		connectString = ini.getValue("AgileAP", "url");

		// 檢查是否Config有空
		if (username.equals("") || password.equals("") || connectString.equals("")) {

			log.log("AgileAP 裏有空！請填完再重新跑程式！");
			displayError();
			System.exit(1);
		}

		try {
			log.log("嘗試登入Agile PLM環境。。");
			session = WebClient.getAgileSession(username, password, connectString);
			if (session != null)
				log.log(1, "登入成功");

		} catch (APIException e) {

			log.log("登入失敗，請確認[AgileAP]的資料");
			displayError();
			System.exit(1);
		}
		testSQLConnection();
		testRunMail();
		
		
		
	}

	/*
	 * 錯誤訊息
	 */
	private static void displayError() {
		JOptionPane.showMessageDialog(null, "程式出錯! 請檢查LOG檔!", "Error Message", JOptionPane.ERROR_MESSAGE);
	}
	
	private static void testSQLConnection(){
		log.log("測試DB的連線");
		Connection conA = AUtil.getDbConn(ini, "AgileDB");
		//random sql string
		String sql = "Select * from activity";
		try {
			log.log("嘗試連接database...");
			conA.createStatement().executeQuery(sql);
			log.log("成功");
		} catch (Exception e) {
			displayError();
			log.log("連接失敗，請檢查AgileDB的設定");
			System.exit(1);
		}
	}

	/*
	 * 主程式 流程 
	 * 1.讀取Excel 
	 * 2.讀取超時工作 
	 * 3.讀取未簽核表單(Change,PSR,QCR) 
	 * 4.寄信給使用者
	 */
	public static void main(String[] args) {

		// 初始化
		Notify notify = new Notify();

		try {
			processExcelforPPM();
			getOverdueProjects();
			log.log("已讀完PPM模組部分，即將開始PC模組的部分\n");
			
			notify.getIncompleteChange();
			notify.getIncompletePSR();
			notify.getIncompleteQCR();
			log.log("讀取結束，準備寄信");

			notifyRelevantUsers(props);
			log.log("程式正常結束,無異常.\n");

		} catch (Exception e) {
			log.log(e);
			displayError();
			System.exit(1);
		}

	}

	private static Properties initializeProperties() {
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
			log.log(e);
			log.log("系統在設定寄件性能時出錯,請確認[Admin Mail]裡的設定都是正確再重試");
			displayError();
			System.exit(1);
		}
		return props;
	}

	private static void testRunMail() {
		// initialize email properties
		String username, password;
		username = ini.getValue("Admin Mail", "username");
		password = ini.getValue("Admin Mail", "password");
		props = initializeProperties();
		Session mailSession = Session.getDefaultInstance(props, null);
		log.log("測試是否能連接到MAIL SERVER");
		try {
			MimeMessage message = new MimeMessage(mailSession);
			Transport transport = mailSession.getTransport();
			Multipart email = new MimeMultipart();
			
			message.setContent(email);
			message.addRecipient(Message.RecipientType.TO, new InternetAddress("william@anselm.com.tw"));
			transport.connect(username, password);
			transport.close();
			log.log("成功連接MAIL");

		} catch (MessagingException e) {
			log.log(e);
			log.log("登入失敗,請查看 [Admin Mail] 的前3項是否正確");
			displayError();
			System.exit(1);
		}

	}

	/*
	 * 寄信給之前所記錄的所有用戶
	 * 郵件的詳細內容都在這裡所產生
	 * @warning 注意事項，logo.png是oracle的logo，需要將檔案存在跟Excel檔同一個資料夾裏
	 */
	private static void notifyRelevantUsers(Properties props) {
		try {
			//寄件者的帳號密碼
			String username = ini.getValue("Admin Mail", "username");
			String password = ini.getValue("Admin Mail", "password");
			Iterator iter = userList.entrySet().iterator();
			log.log("開始發送郵件通知相關人員");
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();
				String key = (String) pair.getKey();
				User usr = (User) pair.getValue();
				String userEmail = usr.getEmail();

				Session mailSession = Session.getDefaultInstance(props, null);
				Transport transport = mailSession.getTransport();
				// 產生整封 email 的主體 message
				MimeMessage message = new MimeMessage(mailSession);

				// 設定主旨
				String subject = "Daily Digest of Updates to your PLM account";
				message.setSubject(subject, "utf-8");
				MimeBodyPart textPart = new MimeBodyPart();
				StringBuffer html = new StringBuffer();
				
				//以下為信件內容

				html.append("<!DOCTYPE html><html><head><style>" + "table,th,td{border: 1px solid black; }"
						+ "td,th{text-align:center;}" + "</style></head><body>");
				html.append("<p>" + pair.getKey() + ", your attention is required</p>");
				html.append("<img src='cid:image'/><br>");

				//檢查用戶是否有新指派的專案
				if (usr.projects.size() != 0) {
					html.append("<p>以下是您被指派的新專案</p>");
					html.append(((User) pair.getValue()).toPPMString());
				}
				
				//檢查用戶是否有未簽核的表單通知
				html.append("<p></p>");
				if (usr.getWorkflow() != null) {
					html.append("<p>以下是您未簽核表單通知</p>");
					html.append(((User) pair.getValue()).toPCString());
				}
				
				//檢查是否有逾期的任務
				html.append("<p></p>");
				if (usr.getOverdueWork() != null) {
					html.append("<p>以下是您的逾期通知</p>");
					html.append(((User) pair.getValue()).toOverdueString());
				}
				html.append("</body></html>");
				html.append("<p>Sincerely,</p><p></p><p>Your Agile PLM Administrator</p>");
				textPart.setContent(html.toString(), "text/html; charset=UTF-8");

				Multipart email = new MimeMultipart();
				email.addBodyPart(textPart);
				MimeBodyPart picturePart = new MimeBodyPart();
				FileDataSource fds = new FileDataSource(ini.getValue("File Location", "FILE_PATH") + "logo.png");
				picturePart.setDataHandler(new DataHandler(fds));
				picturePart.setFileName(fds.getName());
				picturePart.setHeader("Content-ID", "<image>");
				email.addBodyPart(picturePart);
				message.setContent(email);
				//TODO replace with userEmail
				message.addRecipient(Message.RecipientType.TO, new InternetAddress("shsidforever@gmail.com"));
				message.setFrom(new InternetAddress(username));
				transport.connect(username, password);
				log.log("嘗試發送給: " + userEmail);
				try{
					transport.sendMessage(message, message.getRecipients(Message.RecipientType.TO));
					log.log(1,"成功");
				}
				catch (SendFailedException e) {
					log.log(1,"用戶郵箱有問題，跳過");
				}
				transport.close();
			}
		}   catch (MessagingException e){
			log.log(1,"程式在寄信時出錯，請檢查Admin Mail裏面的設定");
			System.exit(1);
			
		}

	}

	/*
	 * 第一部分： 讀取Excel並將相關用戶加進User Class
	 * 由於另一個程式將記錄任何的添加/替換變更，在這裡需要特別再去檢查當下使用者是否符合Excel檔所描述的
	 * @exception APIException 當尋找User時或是project時有可能會出錯
	 * @exception IOException 找不到Excel檔時：1.路徑給錯，2.沒有角色變更
	 */
	private static void processExcelforPPM() {
		try {
			
			Date today = new Date();
			SimpleDateFormat format = new SimpleDateFormat("MMdd");
			filename = format.format(today) + "NoticeChange.xlsx";
			fileLocation = ini.getValue("File Location", "FILE_PATH");
			log.log("讀取Excel檔: "+fileLocation+filename);
			try {
				fis = new FileInputStream(new File(fileLocation + filename));
			} catch (IOException e) {
				log.log("找不到Excel檔案，請檢查File Location里面的FILE_PATH\n跳過");
				return;
			}
			HSSFWorkbook workbook = new HSSFWorkbook(fis);
			HSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();
			// 跳過第一行
			rowIterator.next();
	
			log.log("讀取Excel裡面的資料。。。");
			while (rowIterator.hasNext()) {
				HSSFRow row = (HSSFRow) rowIterator.next();
				Iterator<Cell> cellIterator = row.cellIterator();
				// | User Role	| Name	| user.id |	Email |	Project Name |	Project Link |	Assigned By |
				String role = cellIterator.next().getStringCellValue();
				String name = cellIterator.next().getStringCellValue();
				String userID = cellIterator.next().getStringCellValue();
				String email = cellIterator.next().getStringCellValue();
				String programName = cellIterator.next().getStringCellValue();
				String URL = cellIterator.next().getStringCellValue();
				String assignedBy = cellIterator.next().getStringCellValue();
																				
				log.log("角色: " + role + " 名字: " + name + " ID: " + userID + " 信箱: " + email + " 方案ID: " + programName
						+ " 指派人: " + assignedBy);
				
				// 檢查是否使用者是當下指派的員工
				IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, programName);
				ICell cell = program.getCell("Page Three." + role);
				String userInCell = cell.getValue().toString();
				log.log("檢查使用者是否是當下指派的員工。。	");
	
				// 是 -> 記錄下來
				if (StringUtils.contains(userInCell, userID)) {
					log.log(1, name + " 是當下指派的員工");
					
					//要是程式已經記錄過-> 在已存在的記錄上做變更
					if (userList.get(name) != null) {
						User user = userList.get(name);
						user.addNewProject(programName, URL, assignedBy);
						// adds new role and also check if it is already added
						user.addNewRole(programName, role);
					
					}
					//不然就創新的User然後記錄起來
					else {
						User user = new User(name, userID, email, programName, URL, assignedBy, role);
						userList.put(name, user);
					}
				// 不是-> 跳過
				} else {
					log.log(1, name + " 不是當下指派的員工，這代表在發這封郵件之前他/她曾經有被指派到這個角色但是又被移除了");
					log.log(1, "跳過。");
				}
			}
			fis.close();
		} catch (IOException | APIException e) {
			log.log(e);
			log.log("程式在讀Excel時出錯，請不要亂改程式裏的資料");
			displayError();
		}
	}

	/*
	 * 第二部分： 用SQL找超時的工作
	 * AgileAP裏的url在這時候肯定有值，由於在初始化就有偵測過。
	 */
	public static void getOverdueProjects() {
		try {
			log.log(1, "Accessing database for Overdue tasks");
			//URL是用來之後給用戶鏈接用的
			String URL = "<a href='";
			URL = URL + ini.getValue("AgileAP", "url") + PPM;
			Connection conA = AUtil.getDbConn(ini, "AgileDB");
			String sql = "select name,activity_number,id,sch_end_date,round((sysdate-sch_end_date)) as overdue_duration from activity where act_end_date is null and act_start_date is not null and SYSDATE > sch_end_date order by sch_end_date desc";
			ResultSet rs = conA.createStatement().executeQuery(sql);

			while (rs.next()) {
				String programName = rs.getString(1);
				String programID = rs.getString(2);
				String objID = rs.getString(3);
				Date scheduledEndDate = rs.getDate(4);
				String overdueDuration = rs.getString(5);

				IProgram program = (IProgram) session.getObject(IProgram.OBJECT_TYPE, programID);
				String owner = ((IAgileList) program.getValue(ProgramConstants.ATT_GENERAL_INFO_OWNER)).toString();
				owner = owner.substring(owner.lastIndexOf("(") + 1, owner.lastIndexOf(")"));

				IUser usr = (IUser) session.getObject(IUser.OBJECT_TYPE, owner);
				String firstName = (String) usr.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME);
				String email = (String) usr.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL);
				String type = program.getAgileClass().getAPIName();
				if (userList.containsKey(firstName)) {
					User user = userList.get(firstName);
					String result = userList.get(firstName).getOverdueWork();
					if (result != null) {
						result = result + "<tr><td>" + type + "</td><td nowrap='nowrap'>" + URL + objID + "'>"
								+ programName + "</a></td><td>" + scheduledEndDate + "</td><td>" + overdueDuration
								+ "</td></tr>";
						user.setOverdueWork(result);
						userList.put(firstName, user);
					} else {
						result = "<tr><td>" + type + "</td><td nowrap='nowrap'>" + URL + objID + "'>" + programName
								+ "</a></td><td>" + scheduledEndDate + "</td><td>" + overdueDuration + "</td></tr>";
						user.setOverdueWork(result);
						userList.put(firstName, user);

					}
				} else {
					User user = new User();
					String result = "<tr><td>" + type + "</td><td nowrap='nowrap'>" + URL + objID + "'>" + programName
							+ "</a></td><td>" + scheduledEndDate + "</td><td>" + overdueDuration + "</td></tr>";
					user.setOverdueWork(result);
					user.setEmail(email);
					userList.put(firstName, user);

				}

			}
			conA.close();
		} catch (Exception e) {
			log.log(e);
			displayError();
		}

	}

	/*
	 * 讀取所有未簽核的表單
	 */
	public void getIncompleteChange() throws Exception {
		log.log(1, "Accessing database for Change");
		String URL = "<a href='" + ini.getValue("AgileAP", "url") + SUBADDRESS;
		Connection conA = AUtil.getDbConn(ini, "AgileDB");
		String sql = "select "
				+ "round((sysdate-w.last_upd)-2*FLOOR((sysdate-w.last_upd)/7)-DECODE(SIGN(TO_CHAR(sysdate,'D')-TO_CHAR(w.last_upd,'D')),-1,2,0)+DECODE(TO_CHAR(w.last_upd,'D'),7,1,0)-DECODE(TO_CHAR(sysdate,'D'),7,1,0),2) as WORKDAYS "
				+ ", round((sysdate-w.last_upd),2)  DAYS 			"
				+ ", s.user_name_assigned   		USER_NAME 		"
				+ ", usr.loginid 					USER_ACCOUNT	"
				+ ", n3.description 				CHANGE_TYPE 	"
				+ ", c.change_number 				CHANGE_NUMBER 	"
				+ ", s.last_upd 					LAST_UPD 		"
				+ ", n1.description 				STATUS_FROM		"
				+ ", n2.description 				STATUS_TO 		"
				+ ", c.description 					CHANGE_DESC 	"
				+ ", s.id, s.change_id, s.process_id, s.user_assigned "
				+ "from signoff s, change c, workflow_process w, nodetable n1, nodetable n2, agileuser usr, nodetable n3 "
				+ "where " + "  s.signoff_status=0 and c.delete_flag is null and s.change_id=c.id "
				+ "  and c.process_id=s.process_id and w.id=c.process_id and c.subclass=n3.id "
				+ "and w.state=n1.id and w.next_state=n2.id and usr.id=s.user_assigned "
				+ "order by days desc, s.last_upd desc";
		runSQL(conA, sql, URL);
		log.log("成功讀取資料與database短線中．．．");
		conA.close();

	}

	/*
	 * 讀取所有未簽核的表單
	 */
	public void getIncompleteQCR() throws Exception {

		log.log(1, "Accessing database for QCR");
		String URL = "<a href='" + ini.getValue("AgileAP", "url") + QCR;
		Connection conA = AUtil.getDbConn(ini, "AgileDB");
		String sql = "select "
				+ "round((sysdate-w.last_upd)-2*FLOOR((sysdate-w.last_upd)/7)-DECODE(SIGN(TO_CHAR(sysdate,'D')-TO_CHAR(w.last_upd,'D')),-1,2,0)+DECODE(TO_CHAR(w.last_upd,'D'),7,1,0)-DECODE(TO_CHAR(sysdate,'D'),7,1,0),2) as WORKDAYS "
				+ ", round((sysdate-w.last_upd),2)  DAYS 			"
				+ ", s.user_name_assigned   		USER_NAME 		"
				+ ", usr.loginid 					USER_ACCOUNT	"
				+ ", n3.description 				CHANGE_TYPE 	"
				+ ", c.QCR_NUMBER					CHANGE_NUMBER 	"
				+ ", s.last_upd 					LAST_UPD 		"
				+ ", n1.description 				STATUS_FROM		"
				+ ", n2.description 				STATUS_TO 		"
				+ ", c.description 					CHANGE_DESC 	"
				+ ", s.id, s.change_id, s.process_id, s.user_assigned "
				+ "from signoff s, qcr c, workflow_process w, nodetable n1, nodetable n2, agileuser usr, nodetable n3 "
				+ "where " + "  s.signoff_status=0 and c.delete_flag is null and s.change_id=c.id "
				+ "  and c.process_id=s.process_id and w.id=c.process_id and c.subclass=n3.id "
				+ "and w.state=n1.id and w.next_state=n2.id and usr.id=s.user_assigned "
				+ "order by days desc, s.last_upd desc";
		runSQL(conA, sql, URL);
		log.log("成功讀取資料與database短線中．．．");
		conA.close();

	}

	/*
	 * 讀取所有未簽核的表單
	 */
	public void getIncompletePSR() throws Exception {

		log.log(1, "Accessing database for PSR");
		String URL = "<a href='"+ ini.getValue("AgileAP", "url") + PSR;
		Connection conA = AUtil.getDbConn(ini, "AgileDB");
		String sql = "select "
				+ "round((sysdate-w.last_upd)-2*FLOOR((sysdate-w.last_upd)/7)-DECODE(SIGN(TO_CHAR(sysdate,'D')-TO_CHAR(w.last_upd,'D')),-1,2,0)+DECODE(TO_CHAR(w.last_upd,'D'),7,1,0)-DECODE(TO_CHAR(sysdate,'D'),7,1,0),2) as WORKDAYS "
				+ ", round((sysdate-w.last_upd),2)  DAYS 			"
				+ ", s.user_name_assigned   		USER_NAME 		"
				+ ", usr.loginid 					USER_ACCOUNT	"
				+ ", n3.description 				CHANGE_TYPE 	"
				+ ", c.PSR_NO		 				CHANGE_NUMBER 	"
				+ ", s.last_upd 					LAST_UPD 		"
				+ ", n1.description 				STATUS_FROM		"
				+ ", n2.description 				STATUS_TO 		"
				+ ", c.description 					CHANGE_DESC 	"
				+ ", s.id, s.change_id, s.process_id, s.user_assigned "
				+ "from signoff s, psr c, workflow_process w, nodetable n1, nodetable n2, agileuser usr, nodetable n3 "
				+ "where " + "  s.signoff_status=0 and c.delete_flag is null and s.change_id=c.id "
				+ "  and c.process_id=s.process_id and w.id=c.process_id and c.subclass=n3.id "
				+ "and w.state=n1.id and w.next_state=n2.id and usr.id=s.user_assigned "
				+ "order by days desc, s.last_upd desc";
		runSQL(conA, sql, URL);
		log.log("成功讀取資料與database短線中．．．");
		conA.close();

	}

	/**
	 * @param conA
	 *            used to connect to the database
	 * @param sql
	 *            used to run the SQL
	 * @param URL
	 *            used to create the link in the email.
	 **/
	public void runSQL(Connection conA, String sql, String URL) {
		try {
			ResultSet rs = conA.createStatement().executeQuery(sql);
			while (rs.next()) {
				String userID = rs.getString(4);
				String changeNumber = rs.getString(6);
				String changeType = rs.getString(5);
				String changeDesc = rs.getString(10);
				String status = rs.getString(8);
				String duration = rs.getString(1);
				String changeID = rs.getString(12);
				IUser usr = (IUser) session.getObject(IUser.OBJECT_TYPE, userID);
				String firstName = (String) usr.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME);
				String email = (String) usr.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL);
				log.log(1, "變更類型: " + changeType + " 表單名字: " + changeNumber + " 表單描述: " + changeDesc + " 站別: " + status
						+ " 持續時間: " + duration);
				if (userList.containsKey(firstName)) {

					User user = userList.get(firstName);
					String result = userList.get(firstName).getWorkflow();
					if (result != null) {

						result = result + "<tr><td>" + changeType + "</td><td nowrap='nowrap'>" + URL + changeID + "'>"
								+ changeNumber + "</a></td><td>" + changeDesc + "</td><td>" + status + "</td><td>"
								+ duration + "</td></tr>";
						user.setWorkflow(result);
						userList.put(userID, user);

					} else {
						result = "<tr><td>" + changeType + "</td><td nowrap='nowrap'>" + URL + changeID + "'>"
								+ changeNumber + "</a></td><td>" + changeDesc + "</td><td>" + status + "</td><td>"
								+ duration + "</td></tr>";
						user.setWorkflow(result);
						userList.put(firstName, user);
					}
				} else {
					User user = new User();
					String result = "<tr><td>" + changeType + "</td><td nowrap='nowrap'>" + URL + changeID + "'>"
							+ changeNumber + "</a></td><td>" + changeDesc + "</td><td>" + status + "</td><td>"
							+ duration + "</td></tr>";
					user.setWorkflow(result);
					user.setEmail(email);
					userList.put(firstName, user);
				}
			}
		} catch (Exception e) {
			log.log("執行SQL時遇到問題,請確認 [AgileDB] 的設定正確");
			displayError();
			System.exit(1);
		}
	}

	
}
