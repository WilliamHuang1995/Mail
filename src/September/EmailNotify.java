package September;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
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

import sun.rmi.runtime.Log;

public class EmailNotify extends ServerInfo {
	static IAgileSession m_session;
	IAdmin m_admin;
	AgileSessionFactory m_factory;
	Properties props;
	static HashMap<String, ArrayList<String>> usersMap = new HashMap<String, ArrayList<String>>();
	static final String URL = "<a href='http://192.168.13.250:7001/Agile/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ChangeHandler&classid=6000&objid=";

	public EmailNotify() {
	}

	public static void main(String[] args) {
		try {
			EmailNotify en = new EmailNotify();
			en.run();
		} catch (Exception ex) {
			System.out.println(ex);
		}
	}

	private void run() throws Exception {
		this.m_session = login(username, password, connectString);
		if (m_session != null) {
			System.out.println("Successfully logged in.");
		}

		System.out.println("initializing properties");
		props = initializeProperties();
		System.out.println("properties initialized properly");
		getTable(m_session, null);

	}

	private Properties initializeProperties() {
		Properties props = new Properties();
		try {

			// 初始設定，username 和 password 非必要

			props.setProperty("mail.transport.protocol", "smtp");
			props.setProperty("mail.host", "mail.anselm.com.tw");
			props.setProperty("mail.protocol.port", "25");
			props.setProperty("mail.smtp.auth", "true");

		} catch (Exception e) {
			e.printStackTrace();
			System.exit(1);
		}
		return props;
	}

	private static void sendMail(Properties props) {
		try {

			Iterator iter = usersMap.entrySet().iterator();
			while (iter.hasNext()) {
				Map.Entry pair = (Map.Entry) iter.next();
				// System.out.println(pair.getKey() + " = " + pair.getValue());
				IUser user;
				String key = (String) pair.getKey();
				try {
					user = (IUser) m_session.getObject(UserConstants.CLASS_USER_BASE_CLASS, key);
					String userEmail = (String) user.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL);
					System.out.println(userEmail);
				} catch (APIException e) {
					System.out.println("user not found");
					e.printStackTrace();
				}

				Session mailSession = Session.getDefaultInstance(props, null);
				Transport transport = mailSession.getTransport();
				// 產生整封 email 的主體 message
				MimeMessage message = new MimeMessage(mailSession);

				// 設定主旨
				message.setSubject("您的PLM待簽核表單總覽-2016/09/05");
				MimeBodyPart textPart = new MimeBodyPart();
				StringBuffer html = new StringBuffer();

				//html.append("\n<a href='http://192.168.13.250:7001/Agile/'>Agile PLM</a>");
				html.append("<h3>您的PLM待簽核表單總覽</h3>");
				html.append("<img src='cid:image'/><br>");
				html.append("<table><tr><td>表單編號</td><td>表單描述</td><td>站別</td><td>已持續時間(天)</td></tr>");
				ArrayList<String> val = (ArrayList<String>) pair.getValue();
				while (!val.isEmpty()) {
					html.append(val.get(0));
					val.remove(0);
				}
				html.append("</table>");
				textPart.setContent(html.toString(), "text/html; charset=UTF-8");
				
				//Oracle Logo
				MimeBodyPart picturePart = new MimeBodyPart();
				FileDataSource fds = new FileDataSource("Oracle-logo.png");
				picturePart.setDataHandler(new DataHandler(fds));
				picturePart.setFileName(fds.getName());
				picturePart.setHeader("Content-ID", "<image>");
				
				Multipart email = new MimeMultipart();
				email.addBodyPart(textPart);
				email.addBodyPart(picturePart);
				
				message.setContent(email);
				message.addRecipient(Message.RecipientType.TO, new InternetAddress("johnnylee@anselm.com.tw"));
				message.setFrom(new InternetAddress("william@anselm.com.tw")); // 寄件者
				transport.connect("william@anselm.com.tw", "epacsenur123");
				transport.sendMessage(message, message.getRecipients(Message.RecipientType.TO));
				System.out.println("Completed.");

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

	/*
	 * 郵件內容需包括一個表格，表頭欄位為：表單編號(可超連結)、表單描述、站別、已持續時間(天)
	 */
	public void getTable(IAgileSession session, Map map) throws Exception {

		Ini ini = new Ini("C:/Users/user/Desktop/Anselm/Config.ini");
		LogIt log = new LogIt("");
		Connection conA = null;

		try {
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
			ResultSet rs = conA.createStatement().executeQuery(sql);
			ResultSetMetaData rsmd = rs.getMetaData();
			int numCols = rsmd.getColumnCount();
			// for (int i = 1; i <= numCols; i++)
			// log.log(rsmd.getColumnName(i) + " " + i);
			while (rs.next()) {
				HashMap<String, String> datarow = new HashMap<String, String>();
				String user = rs.getString(4);
				String changeNumber = rs.getString(6);
				String changeDesc = rs.getString(10);
				String status = rs.getString(8);
				String duration = rs.getString(1);
				String changeID = rs.getString(12);
				if (usersMap.containsKey(user)) {
					ArrayList<String> list = usersMap.get(user);
					list.add("<tr><td>" + URL + changeID + "'>" + changeNumber + "</a></td><td>" + changeDesc
							+ "</td><td>" + status + "</td><td>" + duration + "</td></tr>");
					usersMap.put(user, list);
				} else {
					ArrayList<String> list = new ArrayList<String>();
					list.add("<tr><td>" + URL + changeID + "'>" + changeNumber + "</a></td><td>" + changeDesc
							+ "</td><td>" + status + "</td><td>" + duration + "</td></tr>");
					usersMap.put(user, list);
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				conA.close();
				log.log("Agile DB Closed.");
				log.log("Sending Emails.");
				sendMail(props);

			} catch (Exception e) {
			}
		}
	}

	/*
	 * Creating a session and logging in This is sample code provided by Oracle
	 */
	private IAgileSession login(String username, String password, String connectString) throws APIException {

		// Create the params variable to hold login parameters
		HashMap params = new HashMap();

		// Put username and password values into params
		params.put(AgileSessionFactory.USERNAME, username);
		params.put(AgileSessionFactory.PASSWORD, password);

		// Get an Agile server instance. ("agileserver" is the name of the Agile
		// proxy server,
		// and "virtualPath" is the name of the virtual path used for the Agile
		// system.)
		m_factory = AgileSessionFactory.getInstance(connectString);

		// Create the Agile PLM session and log in
		return m_factory.createSession(params);

	}
}
