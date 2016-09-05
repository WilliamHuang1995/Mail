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
			// props.setProperty("mail.user", "william@anselm.com.tw");
			// props.setProperty("mail.password", "epacsenur123");
			props.setProperty("mail.smtp.auth", "true");

			Session mailSession = Session.getDefaultInstance(props, null);
			Transport transport = mailSession.getTransport();

			// 產生整封 email 的主體 message
			MimeMessage message = new MimeMessage(mailSession);

			// 設定主旨
			message.setSubject("Test 9/2/2016");

			// 文字部份，注意 img src 部份要用 cid:接下面附檔的header
			MimeBodyPart textPart = new MimeBodyPart();
			StringBuffer html = new StringBuffer();

			html.append("<h2> Dear Jane </h2><br>");

			// html.append("<h3>這是第二行，下面會是圖</h3><br>");
			html.append("<img src='cid:image'/><br>");
			textPart.setContent(html.toString(), "text/html; charset=UTF-8");

			// 圖檔部份，注意 html 用 cid:image，則header要設<image>
			MimeBodyPart picturePart = new MimeBodyPart();
			FileDataSource fds = new FileDataSource("pokemonbgd.png");
			picturePart.setDataHandler(new DataHandler(fds));
			picturePart.setFileName(fds.getName());
			picturePart.setHeader("Content-ID", "<image>");

			Multipart email = new MimeMultipart();
			email.addBodyPart(textPart);
			email.addBodyPart(picturePart);

			message.setContent(email);
			message.addRecipient(Message.RecipientType.TO, new InternetAddress("paul.huang@foamrite.net"));
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
