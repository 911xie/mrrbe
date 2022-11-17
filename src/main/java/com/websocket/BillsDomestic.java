package com.websocket;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.FileDataSource;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.opslab.helper.HikariAS400;
import com.opslab.util.FileUtil;

/*
 *国外对账 
 */
public class BillsDomestic implements Runnable {

	private static final Logger log = LoggerFactory.getLogger(Service.class);
	private static final String XLS_TEMPLATE = "billsDomestic.xlsx";

	private Service websckt;
	private String sBegin;
	private String sEnd;
	private String sPath;
	private String sEMailAddr;

	public String getsEMailAddr() {
		return sEMailAddr;
	}

	public void setsEMailAddr(String sEMailAddr) {
		this.sEMailAddr = sEMailAddr;
	}

	private String director;

	public String getDirector() {
		return director;
	}

	public void setDirector(String director) {
		this.director = director;
	}

	private String sCode = "";
	private JSONObject rtnObj;

	BillsDomestic() {
		rtnObj = new JSONObject();
		rtnObj.put("code", 0);
		rtnObj.put("msg", "success");
	}

	public String getsCode() {
		return sCode;
	}

	public void setsCode(String sCode) {
		this.sCode = sCode;
	}

	private String sType = "";

	public String getsType() {
		return sType;
	}

	public void setsType(String sType) {
		this.sType = sType;
	}

	private String sFactory = "";

	public String getsFactory() {
		return sFactory;
	}

	public void setsFactory(String sFactory) {
		this.sFactory = sFactory;
	}

	public String getsPath() {
		return sPath;
	}

	public void setsPath(String sPath) {
		this.sPath = sPath;
	}

	private List<String> listTori = new ArrayList<String>(); // 供应商列表

	public String getsBegin() {
		return sBegin;
	}

	public void setsBegin(String sBegin) {
		this.sBegin = sBegin;
	}

	public String getsEnd() {
		return sEnd;
	}

	public void setsEnd(String sEnd) {
		this.sEnd = sEnd;
	}

	public Service getWebsckt() {
		return websckt;
	}

	public void setWebsckt(Service websckt) {
		this.websckt = websckt;
	}

	// 获取所有供应商
	public Map<String, Object> getAllTori() {
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			sCode = sCode.trim().equalsIgnoreCase("ALL") ? "" : " and ASTORI='" + sCode + "' ";
			// 执行sql
//			String sql = "SELECT AA.ASTORIB ASTORI, "
//					+ "case C.ETPATN when NULLIF(C.ETPATN,NULL) THEN C.ETPATN else '' end ETPATN,"
//					+ "case C.ETPOTO when NULLIF(C.ETPOTO,NULL) THEN C.ETPOTO else '' end ETPOTO FROM "
//					+ "(SELECT DISTINCT(ASTORI) ASTORIB "
//					+ "FROM MTKKTAP A INNER JOIN MZKODRP B ON A.ASPONO=B.ZKPONO AND A.ASSEQN=B.ZKSEQN AND B.ZKPLNT='01' "
//					+ "LEFT JOIN MMITEMP C ON A.ASHMCD=C.MIHMCD "
//					+ "LEFT JOIN YMMITMP D ON A.ASHMCD=D.MIHMCD and D.MIPLNT = '999' "
//					+ "WHERE A.ASCNCL<>'D' AND A.ASKDAY BETWEEN " + sBegin + " AND " + sEnd + sCode + ") AA "
//					+ "LEFT JOIN YBMETOLA C ON AA.ASTORIB=C.ETTORI AND C.ETSSKN='AP' "
//					+ "LEFT JOIN YBMTORP D ON AA.ASTORIB=D.MVTORI WHERE MVADCD='D' ORDER BY ASTORIB";
			// 测试物料：BETWEEN 20220919 AND 20220920 AND ASTORI='1091'
			// 物料非ALL
			sType = sType.trim().equalsIgnoreCase("ALL") ? "" : " AND C.MITYPE='" + sType + "' ";

			// 01南屏，02龙山
			// BETWEEN 20220810 AND 20220811 AND A.ASTORI='1812' AND C.MITYPE='15' AND
			// ZKPLNT='01'
			sFactory = sFactory.trim().equalsIgnoreCase("ALL") ? "" : " AND ZKPLNT='" + sFactory + "' ";
			director = director.trim().equalsIgnoreCase("ALL") ? "" : " AND TTJYCD='" + director + "' ";

			String subSql = "SELECT ASTORI, SUM(ASGKIN) AS SUMASGKIN FROM MTKKTAP A "
					+ "INNER JOIN MZKODRP B ON A.ASPONO=B.ZKPONO AND A.ASSEQN=B.ZKSEQN "
					+ "LEFT JOIN MMITEMP C ON A.ASHMCD=C.MIHMCD " + "WHERE ASKDAY BETWEEN " + sBegin + " AND " + sEnd
					+ " AND ZKWFG7<>'Y' " + sCode + sType + sFactory + " GROUP BY ASTORI";

			// 执行sql
			String sql = "SELECT ASTORI, AA.SUMASGKIN, "
					+ "case C.ETPATN when NULLIF(C.ETPATN,NULL) THEN C.ETPATN else '' end ETPATN,"
					+ "case C.ETPOTO when NULLIF(C.ETPOTO,NULL) THEN C.ETPOTO else '' end ETPOTO " + "FROM (" + subSql
					+ ") AA " + "LEFT JOIN YBMETOLA C ON AA.ASTORI=C.ETTORI AND C.ETSSKN='AP' "
					+ "LEFT JOIN YBMTORP D ON AA.ASTORI=D.MVTORI LEFT JOIN UZSKTTP E ON AA.ASTORI=E.TTTORI "
					+ "WHERE AA.SUMASGKIN>0 AND MVADCD='D' " + director + " ORDER BY ASTORI";

			log.info(sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();
			JSONArray jsonAry = new JSONArray();
			Map<String, Object> map = new HashMap<String, Object>();

			log.info("sql=" + sql);
			rs = pstmt.executeQuery();
			while (rs.next()) {
				JSONObject docObj = new JSONObject();
				docObj.put("ASTORI", rs.getString("ASTORI").trim());
				docObj.put("ETPATN", rs.getString("ETPATN").trim());
				docObj.put("ETPOTO", rs.getString("ETPOTO").trim());
				jsonAry.add(docObj);
				// jsonObj.put(rs.getString("ASTORI").trim(), docObj);
				map.put(rs.getString("ASTORI").trim(), docObj);
				listTori.add(rs.getString("ASTORI").trim());
			}
			log.info("map.size=" + map.size());
			jsonObj.put("begin", this.getsBegin());
			jsonObj.put("end", this.getsEnd());
			jsonObj.put("data", jsonAry);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();

			File file = new File(sPath + "record.json");
			FileUtil.write(file, jsonObj.toString(), "UTF-8");

			JSONObject obj = JSONObject.parseObject(FileUtil.read(file, "UTF-8"));
			log.info("aryyyyyyyyyyyy666=" + obj.get("data").toString());
			return map;
		} catch (SQLException e) {
			e.printStackTrace();
		}

		return null;
	}

	// 生成供应商数据
	public JSONObject genToriData(String code, String sBegin, String sEnd) throws Exception {
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		String sql = null;

		rtnObj.put("code", 0);
		rtnObj.put("msg", "success");

		// 创建connection
		conn = HikariAS400.getDataSource().getConnection();
		stmt = conn.createStatement();
		stmt.execute("set current schema MSRFLIB");
		try {
			JSONObject jsonobj = new JSONObject();

			sql = "SELECT A.ASHMCD,A.ASPONO,CASE C.MICNM1 WHEN '' THEN D.MIHMEI ELSE C.MICNM1 END MICNM1,"
					+ "A.ASTANI, A.ASGKAK, A.ASTANK, A.ASKGAK,"
					+ "CASE B.ZKPLNT WHEN '01' THEN '南屏' ELSE '龙山' END ZKPLNT,"
					+ "A.ASKDAY,E.MVTMEK,E.MVBILL FROM MTKKTAP A "
					+ "INNER JOIN MZKODRP B ON A.ASPONO=B.ZKPONO AND A.ASSEQN=B.ZKSEQN "
					+ "LEFT JOIN MMITEMP C ON A.ASHMCD=C.MIHMCD "
					+ "LEFT JOIN YMMITMP D ON A.ASHMCD=D.MIHMCD and D.MIPLNT = '999' "
					+ "LEFT JOIN MMTORIP E ON A.ASTORI = E.MVTORI " + "WHERE A.ASCNCL<>'D' AND A.ASKDAY BETWEEN "
					+ sBegin + " AND " + sEnd + " AND ASTORI='" + code + "'" + sType + sFactory
					+ " ORDER BY A.ASKDAY, B.ZKPLNT";
			log.info("sql=" + sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			String sname = "";
			int tax = 0;
			float sum = 0;
			JSONArray listData = new JSONArray();
			while (rs.next()) {
				JSONObject objData = new JSONObject();
				sname = rs.getString("MVTMEK").trim();
				objData.put("ASHMCD", rs.getString("ASHMCD").trim());// 物品编号
				objData.put("ASPONO", rs.getString("ASPONO").trim());// 订单号
				objData.put("MICNM1", rs.getString("MICNM1").trim());// 物品名称
				objData.put("ASTANI", rs.getString("ASTANI").trim());// 单位
				objData.put("ASGKAK", rs.getInt("ASGKAK"));// 本月检收数量
				objData.put("ASTANK", rs.getFloat("ASTANK"));// 未税单价
				objData.put("ASKGAK", rs.getFloat("ASKGAK"));// 未税金额
				objData.put("ZKPLNT", rs.getString("ZKPLNT").trim());// 检收日期
				objData.put("ASKDAY", rs.getString("ASKDAY"));// 检收日期
				listData.add(objData);

				sum += rs.getFloat("ASKGAK");
				tax = rs.getInt("MVBILL");
			}

			jsonobj.put("ASTORI", code);
			jsonobj.put("MVTMEK", sname);
			jsonobj.put("GenMonth", sBegin.substring(0, 4) + "年" + sBegin.substring(4, 6) + "月");
			jsonobj.put("Tax", tax);
			jsonobj.put("SUM", sum);
			jsonobj.put("SUMWithTax", sum * (1 + tax / 100));

			jsonobj.put("listData", listData);
			log.info("ttttttttttt" + jsonobj.toJSONString());
			fillData(jsonobj, sPath + code + "-" + sBegin + "-" + sEnd + "."
					+ FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE));
		} catch (Exception e) {
			log.info("ddddddddddd" + code + e.getMessage());
			rtnObj.put("code", -1);
			rtnObj.put("msg", e.getMessage());
		} finally {
			rs.close();
			stmt.close();
			conn.close();
		}

		return rtnObj;
	}

	public void fillData(JSONObject jsonobj, String destFile) {
		InputStream in = HikariAS400.class.getClassLoader().getResourceAsStream(BillsDomestic.XLS_TEMPLATE);
		try {
			// InputStream in = new FileInputStream(srcFile);
			FileOutputStream out = new FileOutputStream(destFile);
			log.info("1111111 " + destFile);
			// 读取整个Excel
			XSSFWorkbook wb = new XSSFWorkbook(in);
			// 获取第一个表单Sheet
			XSSFSheet sheetAt = wb.getSheetAt(0);

			int beginLen = 16;
			JSONArray listData = (JSONArray) jsonobj.get("listData");

			// 设置页头内容D8,税率
			sheetAt.getRow(5).getCell(3).setCellValue((String) jsonobj.get("ASTORI"));
			sheetAt.getRow(6).getCell(3).setCellValue((String) jsonobj.get("MVTMEK"));
			sheetAt.getRow(7).getCell(3).setCellValue((Integer) jsonobj.get("Tax"));
			sheetAt.getRow(9).getCell(2).setCellValue((String) jsonobj.get("GenMonth"));
			// 字体
			XSSFFont font = (XSSFFont) wb.createFont();// 创建字体对象
			font.setFontName("微软雅黑");
			font.setFontHeightInPoints((short) 9);

			// 字体
			XSSFFont bigFont = (XSSFFont) wb.createFont();// 创建字体对象
			bigFont.setFontName("微软雅黑");
			bigFont.setFontHeightInPoints((short) 10);
			// 字体
			XSSFFont signFont = (XSSFFont) wb.createFont();// 创建字体对象
			signFont.setFontName("微软雅黑");
			signFont.setFontHeightInPoints((short) 12);

			// XSSFCellStyle保留两位小数
			XSSFCellStyle floatStyle = (XSSFCellStyle) wb.createCellStyle();
			floatStyle.setDataFormat(wb.createDataFormat().getFormat("0.00"));
			floatStyle.setBorderLeft(BorderStyle.THIN);// 左边框
			floatStyle.setBorderBottom(BorderStyle.THIN);// 上边框
			floatStyle.setBorderRight(BorderStyle.THIN);// 右边框
			floatStyle.setFont(font);

			XSSFCellStyle commonStyle = (XSSFCellStyle) wb.createCellStyle();
			commonStyle.setFont(font);
			commonStyle.setBorderLeft(BorderStyle.THIN);// 左边框
			commonStyle.setBorderBottom(BorderStyle.THIN);// 上边框
			commonStyle.setBorderRight(BorderStyle.THIN);// 右边框

			XSSFCellStyle bigFontStyle = (XSSFCellStyle) wb.createCellStyle();
			bigFontStyle.setFont(bigFont);
			bigFontStyle.setBorderLeft(BorderStyle.THIN);// 左边框
			bigFontStyle.setBorderBottom(BorderStyle.THIN);// 上边框
			bigFontStyle.setBorderRight(BorderStyle.THIN);// 右边框

			XSSFCellStyle signStyle = (XSSFCellStyle) wb.createCellStyle();
			signStyle.setFont(signFont);
			signStyle.setAlignment(HorizontalAlignment.CENTER);
			signStyle.setVerticalAlignment(VerticalAlignment.TOP);

			XSSFCellStyle leftBStyle = (XSSFCellStyle) wb.createCellStyle();
			leftBStyle.setBorderLeft(BorderStyle.THIN);
			XSSFCellStyle rightBStyle = (XSSFCellStyle) wb.createCellStyle();
			rightBStyle.setBorderRight(BorderStyle.THIN);

			for (int i = 0; i < listData.size(); i++) {
				int rowNum = beginLen + i;
				XSSFRow rowObj = sheetAt.createRow(rowNum);
				JSONObject itemObj = (JSONObject) listData.get(i);
				for (int j = 1; j < 10; j++) {
					rowObj.createCell(j);
					sheetAt.getRow(rowNum).getCell(j).setCellStyle(commonStyle);
				}

				sheetAt.getRow(rowNum).getCell(6).setCellStyle(floatStyle);
				sheetAt.getRow(rowNum).getCell(7).setCellStyle(floatStyle);
				sheetAt.getRow(rowNum).getCell(1).setCellValue(itemObj.getString("ASHMCD"));
				sheetAt.getRow(rowNum).getCell(2).setCellValue(itemObj.getString("ASPONO"));
				sheetAt.getRow(rowNum).getCell(3).setCellValue(itemObj.getString("MICNM1"));
				sheetAt.getRow(rowNum).getCell(4).setCellValue(itemObj.getString("ASTANI"));
				sheetAt.getRow(rowNum).getCell(5).setCellValue(itemObj.getInteger("ASGKAK"));
				sheetAt.getRow(rowNum).getCell(6).setCellValue(itemObj.getFloat("ASTANK"));
				sheetAt.getRow(rowNum).getCell(7).setCellValue(itemObj.getFloat("ASKGAK"));
				sheetAt.getRow(rowNum).getCell(8).setCellValue(itemObj.getString("ZKPLNT"));
				sheetAt.getRow(rowNum).getCell(9).setCellValue(itemObj.getString("ASKDAY"));
			}

			// 汇总统计
			int currLine = sheetAt.getLastRowNum() + 1;
			XSSFRow rowObj = sheetAt.createRow(currLine);
			for (int j = 1; j < 10; j++) {
				rowObj.createCell(j);
			}
			sheetAt.getRow(currLine).getCell(1).setCellStyle(leftBStyle);
			sheetAt.getRow(currLine).getCell(9).setCellStyle(rightBStyle);
			sheetAt.getRow(currLine).getCell(6).setCellStyle(bigFontStyle);
			sheetAt.getRow(currLine).getCell(6).setCellValue("未税合计");
			sheetAt.getRow(currLine).getCell(7).setCellStyle(floatStyle);
			sheetAt.getRow(currLine).getCell(7)
					.setCellFormula("SUM(H" + beginLen + ":H" + (listData.size() + beginLen) + ")");

			currLine = sheetAt.getLastRowNum() + 1;
			rowObj = sheetAt.createRow(currLine);
			for (int j = 1; j < 10; j++) {
				rowObj.createCell(j);
			}
			sheetAt.getRow(currLine).getCell(1).setCellStyle(leftBStyle);
			sheetAt.getRow(currLine).getCell(9).setCellStyle(rightBStyle);
			sheetAt.getRow(currLine).getCell(6).setCellStyle(bigFontStyle);
			sheetAt.getRow(currLine).getCell(6).setCellValue("含税合计");
			sheetAt.getRow(currLine).getCell(7).setCellStyle(floatStyle);
			int totalLine = listData.size() + beginLen + 1;
			sheetAt.getRow(currLine).getCell(7).setCellFormula("H" + totalLine + "*(1+D8/100)");

			sheetAt.getRow(9).getCell(7).setCellFormula("H" + totalLine);
			sheetAt.getRow(11).getCell(7).setCellFormula("H" + (totalLine + 1));

			// 加空行
			currLine = sheetAt.getLastRowNum() + 1;
			rowObj = sheetAt.createRow(currLine);
			for (int j = 1; j < 10; j++) {
				rowObj.createCell(j);
			}
			sheetAt.getRow(currLine).getCell(1).setCellStyle(leftBStyle);
			sheetAt.getRow(currLine).getCell(9).setCellStyle(rightBStyle);

			// 签章
			currLine = sheetAt.getLastRowNum() + 1;
			CellRangeAddress rangeleft = new CellRangeAddress(currLine, currLine + 5, 1, 3);
			CellRangeAddress rangeright = new CellRangeAddress(currLine, currLine + 5, 4, 9);

			sheetAt.addMergedRegion(rangeleft);
			sheetAt.addMergedRegion(rangeright);

			for (int i = 0; i < 6; i++) {
				XSSFRow rowTmp = sheetAt.createRow(currLine + i);
				for (int j = 1; j < 10; j++) {
					rowTmp.createCell(j);
				}
			}

			sheetAt.getRow(currLine).getCell(1).setCellStyle(signStyle);
			sheetAt.getRow(currLine).getCell(4).setCellStyle(signStyle);
			sheetAt.getRow(currLine).getCell(1).setCellValue("供应商承认/盖章");
			sheetAt.getRow(currLine).getCell(4).setCellValue("紫翔公司承认/盖章");

			// 使用RegionUtil类为合并后的单元格添加边框,这个必须放最后以免样式被覆盖
			RegionUtil.setBorderBottom(BorderStyle.THIN, rangeleft, sheetAt); // 下边框
			RegionUtil.setBorderLeft(BorderStyle.THIN, rangeleft, sheetAt); // 左边框
			RegionUtil.setBorderBottom(BorderStyle.THIN, rangeright, sheetAt); // 下边框
			RegionUtil.setBorderRight(BorderStyle.THIN, rangeright, sheetAt); // 有边框

			sheetAt.setForceFormulaRecalculation(true);

			wb.write(out);
			out.flush();
			out.close();
			in.close();
			wb.close();
		} catch (IOException e) {
			log.info(e.getMessage());
		} finally {

		}

	}

	public static void jacobXls2PDFSample(String inputFilePath, String outputFilePath) {
		ActiveXComponent ax = null;
		Dispatch excel = null;

		try {
			ComThread.InitSTA();
			ax = new ActiveXComponent("Excel.Application");
			ax.setProperty("Visible", new Variant(false));
			// 禁用宏
			ax.setProperty("AutomationSecurity", new Variant(3));
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();
			Object[] obj = { inputFilePath, new Variant(false), new Variant(false) };
			excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

			// 转换格式
			Object[] obj2 = {
					// PDF格式等于0
					new Variant(0), outputFilePath,
					// 0=标准（生成的PDF图片不会模糊），1=最小的文件
					new Variant(0) };
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, obj2, new int[1]);
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {
			if (excel != null) {
				Dispatch.call(excel, "Close", new Variant(false));
			}
			if (ax != null) {
				ax.invoke("Quit", new Variant[] {});
				ax = null;
			}
			ComThread.Release();
		}
	}

	private ActiveXComponent ax = null;
	private Dispatch wb = null;

	public void jacobInit() {
		try {
			ComThread.InitSTA();
			ax = new ActiveXComponent("Excel.Application");
			ax.setProperty("Visible", new Variant(false));
			// 禁用宏
			ax.setProperty("AutomationSecurity", new Variant(3));
			wb = ax.getProperty("Workbooks").toDispatch();
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {

		}
	}

	public void jacobUnInit() {
		if (ax != null) {
			// log.info("2222222222 ax!=null");
			ax.invoke("Quit", new Variant[] {});
			ax = null;
		}
		ComThread.Release();
	}

	public void jacobXls2PDF(String inputFilePath, String outputFilePath) {
		Dispatch excel = null;
		try {
			Object[] obj = { inputFilePath, new Variant(false), new Variant(false) };
			excel = Dispatch.invoke(wb, "Open", Dispatch.Method, obj, new int[9]).toDispatch();

			// 转换格式
			Object[] obj2 = {
					// PDF格式等于0
					new Variant(0), outputFilePath,
					// 0=标准（生成的PDF图片不会模糊），1=最小的文件
					new Variant(0) };
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, obj2, new int[1]);
		} catch (Exception e) {
			e.printStackTrace();
			throw e;
		} finally {
			if (excel != null) {
				Dispatch.call(excel, "Close", new Variant(false));
			}
		}
	}

	public void cleanFin() {
		try {
			websckt.sendMessage("FIN|CLN");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	public void transAndSend(String path) {
		File file = new File(path);
		File[] listFiles = file.listFiles();
		long begin = System.currentTimeMillis();

		File jsonfile = new File(path + "record.json");
		if (!jsonfile.exists()) {
			log.info("error!!!transAndSend json文件不存在。");
			return;
		}

		JSONObject obj = JSONObject.parseObject(FileUtil.read(jsonfile, "UTF-8"));
		String sBegin = obj.get("begin").toString();
		String sEnd = obj.get("end").toString();
		JSONArray objAry = obj.getJSONArray("data");

		int nTotalPdf = 0;
		for (int i = 0; i < listFiles.length; i++) {
			String suffix = FilenameUtils.getExtension(listFiles[i].getName()).toLowerCase();
			File f = listFiles[i];
			if (f.isFile() && f.getName().indexOf("~$") != 0
					&& suffix.equalsIgnoreCase(FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE))) {
				nTotalPdf++;
			}
		}

		log.info("transAndSend jsonfilecnt=" + objAry.size() + ",inpathcnt=" + nTotalPdf);
		if (nTotalPdf != objAry.size()) {
			log.info("error!!!transAndSend 文件数量对不上。");
			return;
		}

		try {
			jacobInit();
			for (int i = 0; i < objAry.size(); i++) {
				JSONObject item = objAry.getJSONObject(i);
				String baseName = item.getString("ASTORI") + "-" + sBegin + "-" + sEnd;
				String srcFile = path + baseName + "." + FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE);
				String destFile = path + baseName + ".pdf";
				File xlsfile = new File(path + baseName + "." + FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE));
				if (xlsfile.exists()) {
					log.info("存在文件:" + baseName);
					jacobXls2PDF(srcFile, destFile);
					if (item.getString("ETPATN").length() > 0) {
						log.info("发送邮件:" + item.getString("ETPATN") + "," + item.getString("ETPOTO"));
						sendMmczMail(item.getString("ETPATN"), item.getString("ETPOTO"), destFile);//
					}
					websckt.sendMessage(String.format("SYN|%02d|生成文件[%s](%d/%d)", ((i + 1) * 100 / nTotalPdf),
							baseName + ".pdf", (i + 1), nTotalPdf));
				}
			}

			jacobUnInit();
			websckt.sendMessage("FIN|2PDF");

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		log.info("spare.1111111111=" + (System.currentTimeMillis() - begin));
	}

	// 邮件参数
	private static class MailInfo {
		String host;
		boolean auth;
		String username;
		String password;
		String from;
		String to;
		String cc;
		String title;
		String content;
		String filename;
	}

	public static void sendMail(MailInfo info) {
		Properties props = new Properties();
		// String host = "MCZ.MMCZ.MEKJPN.NGNET";
		props.put("mail.smtp.host", info.host);// 指定SMTP服务器
		// props.put("mail.smtp.port", String.valueOf(info.port));
		props.put("mail.smtp.auth", String.valueOf(info.auth));// 指定是否需要SMTP验证,为false时，不用指定用户名、密码
		log.info("dddddddd " + String.valueOf(info.auth));

		try {
			Session mailSession = Session.getInstance(props, new Authenticator() {
				@Override
				public PasswordAuthentication getPasswordAuthentication() {
					// 发件人邮箱 用户名和授权码
					return new PasswordAuthentication(info.username, info.password);
				}
			});

			mailSession.setDebug(true);// 是否在控制台显示debug信息

			MimeMessage message = new MimeMessage(mailSession);
			message.setFrom(new InternetAddress(info.from));// 发件人

			// 支持多个收件人
			if (info.to.indexOf(",") != -1) {
				String to2[] = info.to.split(",");
				InternetAddress[] adr = new InternetAddress[to2.length];
				for (int i = 0; i < adr.length; i++) {
					adr[i] = new InternetAddress(to2[i]);
				}
				message.setRecipients(Message.RecipientType.TO, adr);
			} else {
				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(info.to, false));
			}
			if (info.cc != null) {
				if (info.cc.indexOf(",") != -1) {
					// 多个
					String cc2[] = info.cc.split(",");
					InternetAddress[] adr = new InternetAddress[cc2.length];
					for (int i = 0; i < cc2.length; i++) {
						adr[i] = new InternetAddress(cc2[i]);
					}
					// Message的setRecipients方法支持群发。。注意:setRecipients方法是复数和点 到点不一样
					message.setRecipients(Message.RecipientType.CC, adr);
				} else {
					message.setRecipients(Message.RecipientType.CC, InternetAddress.parse(info.cc, false));
				}
			}

			// 将中文转化为GB2312编码
			message.setSubject(info.title, "GB2312"); // 邮件主题
			// message.setContent(info.content, "text/html;charset=gb2312");// 邮件内容
			message.setHeader("Disposition-Notification-To", info.from);

			BodyPart textPart = new MimeBodyPart();
			textPart.setContent(info.content, "text/html;charset=gb2312");

			BodyPart attachPart = new MimeBodyPart();
			FileDataSource fileds = new FileDataSource(info.filename);
			attachPart.setDataHandler(new DataHandler(fileds));
			attachPart.setFileName(fileds.getName());
			Multipart mp = new MimeMultipart();
			mp.addBodyPart(textPart);
			mp.addBodyPart(attachPart);

			message.setContent(mp);
			message.saveChanges();
			Transport transport = mailSession.getTransport("smtp");
			if (info.auth)
				transport.connect(info.host, info.username, info.password);
			else
				transport.connect(info.host, null, null);
			transport.sendMessage(message, message.getAllRecipients());
			// Transport.send(message);
			transport.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void sendMmczMail(String to, String cc, String filename) {
		MailInfo info = new MailInfo();
		info.host = "mmczsmtp.mmcz.mekjpn.ngnet";
		info.auth = false;
		info.username = "";
		info.password = "";
		info.from = this.sEMailAddr;// "xiaojunxie@mektec.com.cn";
		info.to = to;// "xiaojunxie@mektec.com.cn,aibianmu@163.com";
		info.cc = cc;// "911xie@163.com,911xie@126.com";
		info.title = "紫翔对账邮件";
		info.content = "附件是对账文件";
		info.filename = filename;// "D:\\eclipse-workspace\\mrrbe\\tmp\\0001-20220805-20220815.pdf";
		sendMail(info);
	}

	@Override
	public void run() {
		// TODO Auto-generated method stub
		Map<String, Object> map = getAllTori();
		log.info("run..." + map.size() + ":" + map.toString());
		try {
			websckt.sendMessage(String.format("SYN|%02d|正在准备数据...(%d%%)", 5, 5));
			log.info("map.size():" + map.size());

			if (map.size() == 0) {
				websckt.sendMessage("SYN|100|生成供应商数据(0/0)");
				websckt.sendMessage("ERR|查询条件内无供应商数据。");
			} else {

				long begin = System.currentTimeMillis();
				jacobInit();
				for (int i = 0; i < listTori.size(); i++) {
					JSONObject jsonObj = genToriData(listTori.get(i), this.getsBegin(), this.getsEnd());
					log.info("jsonObj=" + jsonObj.toJSONString());
					if (jsonObj.getInteger("code") != 0) {
						websckt.sendMessage("SYN|100|生成供应商数据(0/0)");
						websckt.sendMessage("ERR|" + jsonObj.getString("msg"));
						jacobUnInit();
						return;
					}

					websckt.sendMessage(String.format("SYN|%02d|生成供应商[%s]数据(%d/%d)",
							(5 + (i + 1) * 95 / listTori.size()), listTori.get(i), (i + 1), listTori.size()));
					Thread.sleep(10);
				}
				jacobUnInit();
				websckt.sendMessage("FIN|XLS");
				long end = System.currentTimeMillis();
				log.info("difffffffffff.1=" + (end - begin));
			}

		} catch (Throwable e) {
			log.info("err:" + e.getMessage());
		}

		log.info("文件生成完毕！");
	}

	public static void main(String[] args) throws Exception {
//		BillsDomestic bills = new BillsDomestic();
//		bills.setsPath("D:\\tmp\\");
//		bills.genToriData("1002", "20220901", "20220908");
//
//		bills.jacobInit();
//		bills.jacobXls2PDF("D:\\tmp\\1002-20220901-20220908.xlsx", "D:\\tmp\\1002-20220901-20220908.pdf");
//		bills.jacobUnInit();
		File file = new File("D:\\vueProjects\\mrr\\src\\assets\\images\\background");
		File[] listFiles = file.listFiles();
		for (int i = 0; i < listFiles.length; i++) {
			File f = listFiles[i];
			if (f.isFile()) {
				log.info("'" + f.getName() + "',");
			}
		}
	}

}
