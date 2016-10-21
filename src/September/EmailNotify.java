package September;

import java.io.UnsupportedEncodingException;
import java.sql.Connection;
import java.sql.ResultSet;
import java.util.ArrayList;
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

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAdmin;
import com.agile.api.IAgileSession;
import com.agile.api.IUser;
import com.agile.api.UserConstants;
import com.anselm.plm.util.AUtil;
import com.anselm.plm.utilobj.Ini;
import com.anselm.plm.utilobj.LogIt;
import util.WebClient;

public class EmailNotify {
	static IAgileSession m_session;
	IAdmin m_admin;
	AgileSessionFactory m_factory;
	Properties props;
	static HashMap<String, ArrayList<String>> usersMap = new HashMap<String, ArrayList<String>>();
	static final String SUBADDRESS = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ChangeHandler&classid=6000&objid=";
	static final String PSR = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=PSRHandler&classid=4878&objid=";
	static final String QCR = "/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=QCRHandler&classid=4928&objid=";
	static Ini ini = new Ini();
	static String USERNAME;
	static String PASSWORD;
	static LogIt log = new LogIt("Log");

	public EmailNotify() {
	}

	public static void main(String[] args) {
		try {
			EmailNotify en = new EmailNotify();
			en.run();
		} catch (Exception ex) {
			log.log(ex);
		}
	}

	private void run() throws Exception {
		String filepath = ini.getValue("Settings", "log");
		log.setLogFile(filepath);
		log.log("LOG檔案建立時間: " + new Date());
		log.log("讀取username以及password...");
		String username = ini.getValue("Server Info", "username");
		String password = ini.getValue("Server Info", "password");
		String connectString = ini.getValue("Server Info", "URL");
		log.log(1, "USER: " + username);
		log.log(1, "PASS: " + password);
		log.log(1, "URL : " + connectString);
		log.log("嘗試登入...");
		try {
			EmailNotify.m_session = WebClient.getAgileSession(username, password, connectString);
		} catch (APIException e) {
			log.log(1, "登入失敗, 請確認[Server Info]裡的username, password, URL設定有正確");
			System.exit(1);
		}
		if (m_session != null) {
			log.log(1, "登入成功.");
		}
		props = initializeProperties();
		log.log("讀取 [Settings]");
		// Check to see if needed to send ECO, PSR, QCR
		String ini_change, ini_PSR, ini_QCR;
		ini_change = ini.getValue("Settings", "change");
		ini_PSR = ini.getValue("Settings", "psr");
		ini_QCR = ini.getValue("Settings", "qcr");
		log.log(1, "Change 設定為 " + ini_change);
		log.log(1, "PSR 設定為 " + ini_PSR);
		log.log(1, "QCR 設定為 " + ini_QCR);
		boolean boolChange, boolPSR, boolQCR;
		boolChange = ini_change.equalsIgnoreCase("yes");
		boolPSR = ini_PSR.equalsIgnoreCase("yes");
		boolQCR = ini_QCR.equalsIgnoreCase("yes");

		if (boolChange) {
			log.log("讀取CHANGE相關的未簽核資料");
			getChangeTable(m_session, null);
		}
		if (boolPSR) {
			log.log("讀取PSR相關的未簽核資料");
			getPSRTable(m_session, null);
		}
		if (boolQCR) {
			log.log("讀取QCR相關的未簽核資料");
			getQCRTable(m_session, null);
		}
		if (boolChange || boolPSR || boolQCR) {
			sendMail(props);
		} else {
			log.log("由於Change,QCR,PSR都選擇了no,所以系統不發送郵件");
		}
		log.log("程式正常結束,無異常.\n");
	}

	private Properties initializeProperties() {
		Properties props = new Properties();
		try {
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
			log.log(1, "Email: " + username);
			log.log(1, "Password: " + password);
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
			Iterator iter = usersMap.entrySet().iterator();
			String username = ini.getValue("Admin Mail", "username");
			String password = ini.getValue("Admin Mail", "password");
			testRunMail(username, password, props);
			log.log("開始發送郵件通知相關人員");
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();
				// log.log(pair.getKey() + " = " + pair.getValue());
				IUser user;
				String key = (String) pair.getKey();
				String userEmail = "william@anselm.com.tw";
				try {
					user = (IUser) m_session.getObject(UserConstants.CLASS_USER_BASE_CLASS, key);
					userEmail = (String) user.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL);
					// log.log(userEmail);
				} catch (APIException e) {
					log.log("user not found");
					e.printStackTrace();
				}

				Session mailSession = Session.getDefaultInstance(props, null);
				Transport transport = mailSession.getTransport();
				// 產生整封 email 的主體 message
				MimeMessage message = new MimeMessage(mailSession);

				// 設定主旨
				String subject = "Agile PLM: Changes that still need your approval";
				if (ini.getValue("Admin Mail", "subject") != null) {
					subject = ini.getValue("Admin Mail", "subject");
					try {
						System.out.println(subject);
						subject = new String(subject.getBytes("UTF-8"),"UTF-8");
						System.out.println(subject);
					} catch (UnsupportedEncodingException e) {
						e.printStackTrace();
					}
				}

				message.setSubject(subject, "utf-8");
				MimeBodyPart textPart = new MimeBodyPart();
				StringBuffer html = new StringBuffer();

				html.append("<!DOCTYPE html><html><head><style>" + "table,th,td{border: 1px solid black; }"
						+ "td,th{text-align:center;}" + "h3 {color: maroon;margin-left: 80px;}"
						+ "</style></head><body>");
				html.append("<h3>您的PLM待簽核表單總覽</h3>");
				if (useLogo) {
					html.append("<img src='cid:image'/><br>");
				}
				html.append("<table><tr><td>表單類別</td><td>表單編號</td><td>表單描述</td><td>站別</td><td>已持續時間(天)</td></tr>");
				ArrayList<String> val = (ArrayList<String>) pair.getValue();
				while (!val.isEmpty()) {
					html.append(val.get(0));
					val.remove(0);
				}
				html.append("</table></body></html>");
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
				System.exit(0);
			}
		} catch (AddressException e) {
			e.printStackTrace();
		} catch (NoSuchProviderException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		}

	}

	/*
	 * 郵件內容需包括一個表格，表頭欄位為：表單編號(可超連結)、表單描述、站別、已持續時間(天)
	 */
	public void getChangeTable(IAgileSession session, Map map) throws Exception {
		if (ini.getValue("Server Info", "URL") == null) {
			log.log("請確定 [Server Info] 裡的　URL　有填寫再重新跑一次");
			System.exit(1);
		}
		String URL = "<a href='";
		URL = URL + ini.getValue("Server Info", "URL") + SUBADDRESS;
		Connection conA = null;
		log.log("嘗試連接database...\n要是出錯請確認 database 的 config 是正確的");
		conA = AUtil.getDbConn(ini, "AgileDB");
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

	public void getQCRTable(IAgileSession session, Map map) throws Exception {

		log.log("Accessing database for QCR");
		if (ini.getValue("Server Info", "URL") == null) {
			log.log("請確定 [Server Info] 裡的　URL　有填寫再重新跑一次");
			System.exit(1);
		}
		String URL = "<a href='";
		URL = URL + ini.getValue("Server Info", "URL") + QCR;
		Connection conA = null;
		log.log("嘗試連接database...\n要是出錯請確認 database 的 config 是正確的");
		conA = AUtil.getDbConn(ini, "AgileDB");
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

	public void getPSRTable(IAgileSession session, Map map) throws Exception {

		log.log("Accessing database for PSR");
		if (ini.getValue("Server Info", "URL") == null) {
			log.log("請確定 [Server Info] 裡的　URL　有填寫再重新跑一次");
			System.exit(1);
		}
		String URL = "<a href='";
		URL = URL + ini.getValue("Server Info", "URL") + PSR;
		Connection conA = null;
		log.log("嘗試連接database...\n要是出錯請確認 database 的 config 是正確的");
		conA = AUtil.getDbConn(ini, "AgileDB");
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

	public void runSQL(Connection conA, String sql, String URL) {
		try {
			ResultSet rs = conA.createStatement().executeQuery(sql);
			while (rs.next()) {
				String user = rs.getString(4);
				String changeNumber = rs.getString(6);
				String changeType = rs.getString(5);
				String changeDesc = rs.getString(10);
				String status = rs.getString(8);
				String duration = rs.getString(1);
				String changeID = rs.getString(12);
				if (usersMap.containsKey(user)) {
					ArrayList<String> list = usersMap.get(user);
					list.add("<tr><td>" + changeType + "</td><td nowrap='nowrap'>" + URL + changeID + "'>"
							+ changeNumber + "</a></td><td>" + changeDesc + "</td><td>" + status + "</td><td>"
							+ duration + "</td></tr>");
					usersMap.put(user, list);
				} else {
					ArrayList<String> list = new ArrayList<String>();
					list.add("<tr><td>" + changeType + "</td><td nowrap='nowrap'>" + URL + changeID + "'>"
							+ changeNumber + "</a></td><td>" + changeDesc + "</td><td>" + status + "</td><td>"
							+ duration + "</td></tr>");
					usersMap.put(user, list);
				}
			}
		} catch (Exception e) {
			log.log("執行SQL時遇到問題,請確認 [AgileDB] 的設定正確");
			System.exit(1);
		}
	}
}
