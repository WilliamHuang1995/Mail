package September;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;

public class SendMailDemo {

	/**
	 * Java：以javamail寄送附件圖檔與html格式email教學
	 *
	 * @author werdna at http://werdna1222coldcodes.blogspot.com/
	 */
	public static void main(String[] args) {

		try {

			// 初始設定，username 和 password 非必要
			Properties props = new Properties();
			props.setProperty("mail.transport.protocol", "smtp");
			props.setProperty("mail.host", "mail.anselm.com.tw");
			props.setProperty("mail.protocol.port", "25");
			props.setProperty("mail.smtp.auth", "true");

			Session mailSession = Session.getDefaultInstance(props, null);
			Transport transport = mailSession.getTransport();

			// 產生整封 email 的主體 message
			MimeMessage message = new MimeMessage(mailSession);

			// 設定主旨
			message.setSubject("您的PLM待簽核表單總覽-2016/09/05");
			
			// 文字部份，注意 img src 部份要用 cid:接下面附檔的header
			MimeBodyPart textPart = new MimeBodyPart();
			StringBuffer html = new StringBuffer();
			html.append("<h3>您的PLM待簽核表單總覽</h3>");
			html.append("<img src='cid:image'/><br>");	
			html.append("<table><tr><td>表單編號</td><td>表單描述</td><td>站別</td><td>已持續時間(天)</td></tr>");
			html.append("<tr><td>1</td><td>2</td><td>3</td><td>4</td></tr>");
			html.append("<tr><td>1</td><td>2</td><td>3</td><td>4</td></tr>");
			html.append("<tr><td><a href='http://192.168.13.250:7001/Agile/PCMServlet?fromPCClient=true&module=ChangeHandler&requestUrl=module%3DChangeHandler%26opcode%3DdisplayObject%26classid%3D6000%26objid%3D7153413%26tabid%3D0%26'>test</a></td><td><a href='http://192.168.13.250:7001/Agile/PLMServlet?action=OpenEmailObject&isFromNotf=true&module=ChangeHandler&classid=6000&objid=7153413'>test2</a></td><td>3</td><td>4</td></tr></table>");
			html.append("<p><a href='http://192.168.13.250:7001/Agile/PCMServlet?fromPCClient=true&module=ChangeHandler&requestUrl=module%3DChangeHandler%26opcode%3DdisplayObject%26classid%3D6000%26objid%3D7153413%26tabid%3D0%26'>點取這裡可直接進入客戶端</a></p>");
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
			message.addRecipient(Message.RecipientType.TO, new InternetAddress("wjhuang@ucsd.edu"));
			message.setFrom(new InternetAddress("william@anselm.com.tw")); // 寄件者
			transport.connect("william@anselm.com.tw", "epacsenur123");
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
