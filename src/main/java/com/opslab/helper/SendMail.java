package com.opslab.helper;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;

import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
public class SendMail {

	private static final Logger log = LoggerFactory.getLogger(SendMail.class);

	private String host;
	private boolean auth;
	private String username;
	private String password;
	private String from;
	private String to;
	private String cc;
	private String title;
	private String content;
	// private String attachFile = "";
	private JSONArray attachFile = new JSONArray();
	private String imageFile = "";

	public void sendMail() {
		Properties props = new Properties();
		// String host = "MCZ.MMCZ.MEKJPN.NGNET";
		props.put("mail.smtp.host", host);// 指定SMTP服务器
		// props.put("mail.smtp.port", String.valueOf(port));
		props.put("mail.smtp.auth", String.valueOf(auth));// 指定是否需要SMTP验证,为false时，不用指定用户名、密码

		try {
			// Session mailSession = Session.getDefaultInstance(props);

			// 这一步是qq邮箱才有, 其他邮箱不用
//			if (host.indexOf("smtp.qq.com") >= 0) {
//				MailSSLSocketFactory sf = new MailSSLSocketFactory();
//				sf.setTrustAllHosts(true);
//				props.put("mail.smtp.ssl.enable", "true");
//				props.put("mail.smtp.ssl.socketFactory", sf);
//				log.info("eeeeee " + host);
//			}

			Session mailSession = Session.getInstance(props, new Authenticator() {
				@Override
				public PasswordAuthentication getPasswordAuthentication() {
					// 发件人邮箱 用户名和授权码
					return new PasswordAuthentication(username, password);
				}
			});

			// mailSession.setDebug(true);// 是否在控制台显示debug信息

			MimeMessage message = new MimeMessage(mailSession);
			message.setFrom(new InternetAddress(from));// 发件人

//			message.addRecipient(Message.RecipientType.TO, new InternetAddress(to));// 收件人
//			message.setRecipient(Message.RecipientType.CC, new InternetAddress(cc));// 收件人

			// 支持多个收件人
			if (to.indexOf(",") != -1) {
				String to2[] = to.split(",");
				InternetAddress[] adr = new InternetAddress[to2.length];
				for (int i = 0; i < adr.length; i++) {
					adr[i] = new InternetAddress(to2[i]);
				}
				message.setRecipients(Message.RecipientType.TO, adr);
			} else {
				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(to, false));
			}
			if (cc != null) {
				if (cc.indexOf(",") != -1) {
					// 多个
					String cc2[] = cc.split(",");
					InternetAddress[] adr = new InternetAddress[cc2.length];
					for (int i = 0; i < cc2.length; i++) {
						adr[i] = new InternetAddress(cc2[i]);
					}
					// Message的setRecipients方法支持群发。。注意:setRecipients方法是复数和点 到点不一样
					message.setRecipients(Message.RecipientType.CC, adr);
				} else {
					message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(cc, false));
				}
			}

			// 将中文转化为GB2312编码
			message.setSubject(title, "GB2312"); // 邮件主题
			// message.setContent(content, "text/html;charset=gb2312");// 邮件内容
			message.setHeader("Disposition-Notification-To", from);

			MimeMultipart multipart = new MimeMultipart("related");

			// first part (the html)
			BodyPart bp = new MimeBodyPart();
			String htmlText = imageFile.length() > 0 ? content + "<br><img src='cid:image'>" : content;
			bp.setContent(htmlText, "text/html;charset=gb2312");
			// add it
			multipart.addBodyPart(bp);

			// second part (the image)
			if (imageFile.length() > 0) {
				bp = new MimeBodyPart();
				DataSource fds = new FileDataSource(imageFile);
				bp.setDataHandler(new DataHandler(fds));
				bp.setHeader("Content-ID", "<image>");
				multipart.addBodyPart(bp);
			}

			for (int i = 0; i < attachFile.size(); i++) {
				bp = new MimeBodyPart();
				FileDataSource fileds = new FileDataSource(attachFile.getString(i));
				bp.setDataHandler(new DataHandler(fileds));
				bp.setFileName(fileds.getName());
				multipart.addBodyPart(bp);
			}

			message.setContent(multipart);
			message.saveChanges();
			Transport transport = mailSession.getTransport("smtp");
			if (auth)
				transport.connect(host, username, password);
			else
				transport.connect(host, null, null);
			transport.sendMessage(message, message.getAllRecipients());
			// Transport.send(message);
			transport.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) throws Exception {
		SendMail sendMail = new SendMail();
		sendMail.host = "mmczsmtp.mmcz.mekjpn.ngnet";
		sendMail.auth = false;
		sendMail.username = "";
		sendMail.password = "";
		sendMail.from = "MPE_Report@mektec.com.cn";
		sendMail.to = "xiaojunxie@mektec.com.cn";
		sendMail.cc = "911xie@163.com";
		sendMail.title = "紫翔对账邮件";
//		sendMail.content = "<table width=\"703\" style=\"border-collapse:collapse;\">\r\n" + "	<tr height=\"8\">\r\n"
//				+ "		<td width=\"90\" bgcolor=\"#6fa1d2\" style=\"border-style:none none solid none;border-color:#5C87B2;border-width:0px 0px 1px 0px;padding:1px 1px;\">\r\n"
//				+ "			<div align=\"center\"><span style=\" font-size:12pt;color:white\">Date</span></div>\r\n"
//				+ "		</td>\r\n"
//				+ "		<td width=\"105\" bgcolor=\"#6fa1d2\" style=\"border-style:none none solid none;border-color:#5C87B2;border-width:0px 0px 1px 0px;padding:1px 1px;\">\r\n"
//				+ "			<div align=\"center\"><span style=\" font-size:12pt;color:white\">Assy MPE Rate</span></div>\r\n"
//				+ "		</td>\r\n" + "	</tr>\r\n" + "		<tr height=\"8\">\r\n"
//				+ "			<td width=\"89\" style=\"border-style:solid solid solid solid;border-color:#5C87B2;border-width:1px 1px 1px 1px;padding:1px 1px;\">\r\n"
//				+ "				<div align=\"center\"><span style=\" font-size:12pt;color:#4f4f4f\">10-10-10-16</span></div>\r\n"
//				+ "			</td>\r\n"
//				+ "			<td width=\"104\" style=\"border-style:solid solid solid solid;border-color:#5C87B2;border-width:1px 1px 1px 1px;padding:1px 1px;\">\r\n"
//				+ "				<div align=\"center\"><span style=\" font-size:12pt;color:#4f4f4f\">0%</span></div>\r\n"
//				+ "			</td>\r\n" + "	</tr>\r\n" + "</table>";// "附件是对账文件";
		sendMail.content = "附件是对账文件1";
		sendMail.imageFile = "D:\\mpechart.png";
		JSONArray ary = new JSONArray();
		ary.add("D:\\eclipse-workspace\\mrrbe\\Reconcile.xltx");
		ary.add("D:\\eclipse-workspace\\mrrbe\\Reconcile3.xltx");
		sendMail.attachFile = ary;
		sendMail.sendMail();

	}

}
