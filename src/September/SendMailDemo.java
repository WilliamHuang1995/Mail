package September;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;

import com.anselm.plm.utilobj.Ini;

public class SendMailDemo {

	/**
	 * Java：以javamail寄送附件圖檔與html格式email教學
	 *
	 */
	public static void main(String[] args) {

		try {
			Ini ini = new Ini();
			String path = ini.getValue("Server Info", "URL");
			System.out.println(path);
			// 初始設定，username 和 password 非必要
			Properties props = new Properties();
			String transportProtocol = ini.getValue("Admin Mail", "transport protocol");
			String host = ini.getValue("Admin Mail", "host");
			String protocolPort = ini.getValue("Admin Mail", "protocol port");
			props.setProperty("mail.transport.protocol", transportProtocol);
			props.setProperty("mail.host", host);
			props.setProperty("mail.protocol.port", protocolPort);
			props.setProperty("mail.smtp.auth", "true");

			Session mailSession = Session.getDefaultInstance(props, null);
			Transport transport = mailSession.getTransport();

			// 產生整封 email 的主體 message
			MimeMessage message = new MimeMessage(mailSession);

			// 設定主旨
			message.setSubject("Agile PLM: Changes that still need your approval");
			
			// 文字部份，注意 img src 部份要用 cid:接下面附檔的header
			MimeBodyPart textPart = new MimeBodyPart();
			StringBuffer html = new StringBuffer();
			html.append("<!DOCTYPE html><html><head><style>table,th,td{border: 1px solid black; }td,th{text-align:center;}h3 {    color: maroon;    margin-left: 80px;}</style></head><body>");
			html.append("<h3>您的PLM待簽核表單總覽</h3>");
			html.append("<img src='cid:image'/><br>");	
			html.append("<table><tr><th>表單編號</th><th>表單描述</th><th>站別</th><th>已持續時間(天)</th></tr>");
			html.append("<tr><td>1</td><td>2</td><td>3</td><td>4</td></tr>");
			html.append("<tr><td>1</td><td>2</td><td>3</td><td>4</td></tr>");
			html.append("<tr><td><a href='http://192.168.13.250:7001/Agile/PCMServlet?fromPCClient=true&module=ChangeHandler&requestUrl=module%3DChangeHandler%26opcode%3DdisplayObject%26classid%3D6000%26objid%3D7153413%26tabid%3D0%26'>test</a></td><td><a href='http://192.168.13.250:7001/Agile/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ChangeHandler&classid=6000&objid=7153413'>test2</a></td><td>3</td><td>4</td></tr></table>");
			html.append("<p><a href='http://192.168.13.250:7001/Agile/PCMServlet?fromPCClient=true&module=ChangeHandler&requestUrl=module%3DChangeHandler%26opcode%3DdisplayObject%26classid%3D6000%26objid%3D7153413%26tabid%3D0%26'>點取這裡可直接進入客戶端</a></p>");
			html.append("</body></html>");
			textPart.setContent(html.toString(), "text/html; charset=UTF-8"); 

			// 圖檔部份，注意 html 用 cid:image，則header要設<image>
			MimeBodyPart picturePart = new MimeBodyPart();
			FileDataSource fds = new FileDataSource("Oracle-logo.png");
			picturePart.setDataHandler(new DataHandler(fds));
			picturePart.setFileName(fds.getName());
			picturePart.setHeader("Content-ID", "<image>");

			Multipart email = new MimeMultipart();
			email.addBodyPart(textPart);
			email.addBodyPart(picturePart);

			message.setContent(email);
			//replace wjhuang@ucsd.edu with userEmail
			message.addRecipient(Message.RecipientType.TO, new InternetAddress("william@anselm.com.tw"));
			String username = ini.getValue("Admin Mail", "username");
			String password = ini.getValue("Admin Mail", "password");
			message.setFrom(new InternetAddress(username)); // 寄件者
			transport.connect(username, password);
			transport.sendMessage(message, message.getRecipients(Message.RecipientType.TO));
			System.out.println("Completed.");
			transport.close();
		} catch (AddressException e) {
			e.printStackTrace();
		} catch (NoSuchProviderException e) {
			e.printStackTrace();
		} catch (MessagingException e) {
			e.printStackTrace();
		}
	}
}
