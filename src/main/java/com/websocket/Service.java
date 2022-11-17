package com.websocket;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;

import javax.websocket.EncodeException;
import javax.websocket.OnClose;
import javax.websocket.OnError;
import javax.websocket.OnMessage;
import javax.websocket.OnOpen;
import javax.websocket.Session;
import javax.websocket.server.PathParam;
import javax.websocket.server.ServerEndpoint;

import org.apache.commons.lang.StringUtils;
import org.springframework.stereotype.Component;

import com.alibaba.fastjson.JSONObject;
import com.opslab.helper.HikariMySql;

import lombok.extern.slf4j.Slf4j;

//@ServerEndpoint(value = "/websocket/{userName}")
@ServerEndpoint(value = "/bills/{userName}")
@Component
@Slf4j
public class Service {

	// private static final Logger log = LoggerFactory.getLogger(Service.class);

	// 静态变量，用来记录当前在线连接数。应该把它设计成线程安全的。
	private static int onlineCount = 0;
	// concurrent包的线程安全Set，用来存放每个客户端对应的WebSocketServer对象。
	private static ConcurrentHashMap<String, Client> webSocketMap = new ConcurrentHashMap<>();

	/** 与某个客户端的连接会话，需要通过它来给客户端发送数据 */
	private Session session;
	/** 接收userName */
	private String userName = "";

	private String outPutPath = "";
	private String sendMail = "";

	public Service() {
		InputStream in = Service.class.getClassLoader().getResourceAsStream("baseconfig.properties");
		Properties properties = new Properties();

		try {
			// properties.load(in);
			properties.load(new InputStreamReader(in, "UTF-8"));
			in.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		outPutPath = properties.getProperty("baseconfig.outPutPath");
		sendMail = properties.getProperty("baseconfig.sendMail");
		log.info("1.configfile read=" + outPutPath + "," + sendMail);
		getConfigFromDb();
	}

	private void getConfigFromDb() {
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;

		try {
			conn = HikariMySql.getDataSource().getConnection();
			String sql = "select * from config";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			JSONObject jsonobj = new JSONObject();
			while (rs.next()) {
				jsonobj.put(rs.getString("key").trim(), rs.getString("value").trim());
			}

			outPutPath = jsonobj.getString("outPutPath").trim();
			sendMail = jsonobj.getString("sendMail").trim();

			rs.close();
			pstmt.close();
			conn.close();
			log.info("22.db read=" + outPutPath + "," + sendMail);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	/**
	 * 连接建立成功调用的方法
	 */
	@OnOpen
	public void onOpen(Session session, @PathParam("userName") String userName) {
		if (!webSocketMap.containsKey(userName)) {
			addOnlineCount(); // 在线数 +1
		}
		this.session = session;
		this.userName = userName;
		Client client = new Client();
		client.setSession(session);
		client.setUri(session.getRequestURI().toString());
		webSocketMap.put(userName, client);

		log.info("----------------------------------------------------------------------------");
		log.info("用户连接:" + userName + ",当前在线人数为:" + getOnlineCount());

		try {
			sendMessage("来自后台的反馈：连接成功");
		} catch (IOException e) {
			log.error("用户:" + userName + ",网络异常!!!!!!");
		}
	}

	/**
	 * 连接关闭调用的方法
	 */
	@OnClose
	public void onClose() {
		if (webSocketMap.containsKey(userName)) {
			webSocketMap.remove(userName);
			if (webSocketMap.size() > 0) {
				// 从set中删除
				subOnlineCount();
			}
		}
		log.info("----------------------------------------------------------------------------");
		log.info(userName + "用户退出,当前在线人数为:" + getOnlineCount());
	}

	public static void createFile(String path) {
		InputStream in = null;
		OutputStream out = null;
		try {
//			File srcFile = new File("Z:\\ZP188006\\188003.pdf");
//			File destFile = new File("Z:\\ZP188006\\dest188003.pdf");
			File srcFile = new File(path + "188003.pdf");
			File destFile = new File(path + "dest777-188003.pdf");
			in = new BufferedInputStream(new FileInputStream(srcFile));
			out = new BufferedOutputStream(new FileOutputStream(destFile));
			byte[] buffer = new byte[4096];
			int len = 0; // 读取长度
			while ((len = in.read(buffer, 0, buffer.length)) != -1) {
				out.write(buffer, 0, len);
			}
			out.flush();// 刷新缓冲的输出流
		} catch (Exception e) {
			log.info("读取文件发生错误！");
		} finally {
			try {
				if (out != null) {
					out.close();
				}
				if (in != null) {
					in.close();
				}
			} catch (Exception e) {
				log.info("关闭资源发生错误！");
			}
		}
	}

	// 删除指定文件夹下所有文件
	// param path 文件夹完整绝对路径
	public static boolean delAllFile(String path) {
		boolean flag = false;
		File file = new File(path);
		if (!file.exists()) {
			return flag;
		}
		if (!file.isDirectory()) {
			return flag;
		}
		String[] tempList = file.list();
		File temp = null;
		for (int i = 0; i < tempList.length; i++) {
			if (path.endsWith(File.separator)) {
				temp = new File(path + tempList[i]);
			} else {
				temp = new File(path + File.separator + tempList[i]);
			}
			if (temp.isFile()) {
				temp.delete();
			}
			if (temp.isDirectory()) {
				delAllFile(path + "/" + tempList[i]);// 先删除文件夹里面的文件
//              delFolder(path + "/" + tempList[i]);// 再删除空文件夹
				flag = true;
			}
		}
		return flag;
	}

	public void onMessageAbroad(String[] ary) {
		// String path = "D:\\eclipse-workspace\\mrrbe\\tmp\\";
		String path = this.outPutPath + this.userName + "\\Abroad\\";
		// 判断文件路径是否存在不存在创建新的文件
		File file = new File(path);
		if (!file.exists()) {
			file.mkdirs();
		}

		BillsAbroad work = new BillsAbroad();
		work.setWebsckt(this);
		work.setDirector(userName);
		work.setsPath(path);
		work.setsEMailAddr(this.sendMail);

		switch (ary[1]) {
		case "Bills":
			log.info("收到用户Bills命令" + ary[2] + ary[3] + ary[4] + ary[5] + ary[6]);
			String toriCode = "";

			work.setsBegin(ary[2].replaceAll("-", ""));
			work.setsEnd(ary[3].replaceAll("-", ""));
			work.setsCode(ary[4]);
			work.setsType(ary[5]);
			work.setsFactory(ary[6]);
			work.run();
			break;
		case "TransPDFandSend":
			log.info("收到用户TransPDFandSend命令");
			work.transAndSend(path);
			break;
		case "Clean":
			log.info("收到用户Clean命令" + path);
			Service.delAllFile(path);
			work.cleanFin();
			break;
		default:
			log.info("unknow cmd!");
			break;
		}
	}

	public void onMessageDomestic(String[] ary) {
		// String path = "D:\\eclipse-workspace\\mrrbe\\tmp\\";
		String path = this.outPutPath + this.userName + "\\Domestic\\";
		// 判断文件路径是否存在不存在创建新的文件
		// path = CharsetHepler.toUTF_8(path);
		File file = new File(path);
		if (!file.exists()) {
			file.mkdirs();
		}

		BillsDomestic work = new BillsDomestic();
		work.setWebsckt(this);
		work.setDirector(userName);
		work.setsPath(path);
		work.setsEMailAddr(this.sendMail);

		switch (ary[1]) {
		case "Bills":
			log.info("收到用户Bills命令" + ary[2] + ary[3] + ary[4] + ary[5] + ary[6]);
			String toriCode = "";
			work.setsBegin(ary[2].replaceAll("-", ""));
			work.setsEnd(ary[3].replaceAll("-", ""));
			work.setsCode(ary[4]);
			work.setsType(ary[5]);
			work.setsFactory(ary[6]);
			work.run();
			break;
		case "TransPDFandSend":
			log.info("收到用户TransPDFandSend命令");
			work.transAndSend(path);
			break;
		case "Clean":
			log.info("收到用户Clean命令" + path);
			Service.delAllFile(path);
			work.cleanFin();
			break;
		default:
			log.info("unknow cmd!");
			break;
		}
	}

	/**
	 * 收到客户端消息后调用的方法
	 *
	 * @param message 客户端发送过来的消息
	 * @throws InterruptedException
	 */
	@OnMessage
	public void onMessage(String message, Session session) throws InterruptedException {
		log.info("收到用户" + userName + "消息,报文:" + message);

		// log.info("config.2=" + config.getSendMail() + "," + config.getOutPutPath());
		// 可以群发消息
		// 消息保存到数据库、redis
		if (StringUtils.isNotBlank(message)) {
			// "|"是特殊字符，必须加转义\\|才行
			String[] ary = message.split("\\|");
			switch (ary[0]) {
			case "Abroad":
				onMessageAbroad(ary);
				break;
			case "Domestic":
				onMessageDomestic(ary);
				break;
			default:
				log.info("unknow cmd!");
				break;
			}
		}
	}

	/**
	 *
	 * @param session
	 * @param error
	 */
	@OnError
	public void onError(Session session, Throwable error) {
		log.error("用户错误:" + this.userName + ",原因:" + error.getMessage());
		error.printStackTrace();
	}

	/**
	 * 连接服务器成功后主动推送
	 */
	public void sendMessage(String message) throws IOException {
		synchronized (session) {
			this.session.getBasicRemote().sendText(message);
		}
	}

	public void sendMessage(Object data) throws IOException, EncodeException {
		synchronized (session) {
			this.session.getBasicRemote().sendObject(data);
		}
	}

	/**
	 * 向指定客户端发送消息
	 * 
	 * @param userName
	 * @param message
	 */
	public static void sendMessage(String userName, String message) {
		try {
			Client webSocketClient = webSocketMap.get(userName);
			if (webSocketClient != null) {
				webSocketClient.getSession().getBasicRemote().sendText(message);
			}
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException(e.getMessage());
		}
	}

	public static synchronized int getOnlineCount() {
		return onlineCount;
	}

	public static synchronized void addOnlineCount() {
		Service.onlineCount++;
	}

	public static synchronized void subOnlineCount() {
		Service.onlineCount--;
	}

	public static void setOnlineCount(int onlineCount) {
		Service.onlineCount = onlineCount;
	}

	public static ConcurrentHashMap<String, Client> getWebSocketMap() {
		return webSocketMap;
	}

	public static void setWebSocketMap(ConcurrentHashMap<String, Client> webSocketMap) {
		Service.webSocketMap = webSocketMap;
	}

	public Session getSession() {
		return session;
	}

	public void setSession(Session session) {
		this.session = session;
	}

	public String getUserName() {
		return userName;
	}

	public void setUserName(String userName) {
		this.userName = userName;
	}

}
