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
import java.util.Iterator;
import java.util.List;
import java.util.Map;

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
import com.opslab.helper.SendMail;
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
	private String factory = "";

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
		factory = sFactory.trim();
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
			String sql = "SELECT DISTINCT ASTORI, AA.SUMASGKIN, "
					+ "case C.ETPATN when NULLIF(C.ETPATN,NULL) THEN C.ETPATN else '' end ETPATN,"
					+ "case C.ETPOTO when NULLIF(C.ETPOTO,NULL) THEN C.ETPOTO else '' end ETPOTO " + "FROM (" + subSql
					+ ") AA " + "LEFT JOIN YBMETOLA C ON AA.ASTORI=C.ETTORI AND C.ETSSKN='AP' AND ETATNO=1 "
					+ "LEFT JOIN YBMTORP D ON AA.ASTORI=D.MVTORI LEFT JOIN UZSKTTP E ON AA.ASTORI=E.TTTORI "
					+ "WHERE AA.SUMASGKIN>0 AND MVADCD='D' " + director + " ORDER BY ASTORI";

			log.info(sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();

			Map<String, Object> map = new HashMap<String, Object>();

			log.info("sql=" + sql);
			rs = pstmt.executeQuery();

			File file = new File(sPath + "record.json");
			JSONObject fileObj = JSONObject.parseObject(FileUtil.read(file, "UTF-8"));
			if (fileObj == null) {
				fileObj = new JSONObject();
			}

			// JSONObject docASTORI = new JSONObject();
			while (rs.next()) {
				JSONObject docObj = new JSONObject();
				docObj.put("ASTORI", rs.getString("ASTORI").trim());
				docObj.put("ETPATN", rs.getString("ETPATN").trim());
				docObj.put("ETPOTO", rs.getString("ETPOTO").trim());
				docObj.put("SEND", 0);

				String fileName = rs.getString("ASTORI").trim() + "-" + factory + "-" + this.getsBegin() + "-"
						+ this.getsEnd();// + "." + FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE);

				if (!fileObj.containsKey(rs.getString("ASTORI").trim())) {
					JSONArray jsonAry = new JSONArray();
					jsonAry.add(fileName);
					docObj.put("FileNames", jsonAry);
					fileObj.put(rs.getString("ASTORI").trim(), docObj);
				} else {
					JSONObject fnObj = (JSONObject) fileObj.get(rs.getString("ASTORI").trim());
					JSONArray jsonAry = fnObj.getJSONArray("FileNames");
					jsonAry.add(fileName);
					docObj.put("FileNames", jsonAry);
				}

				// jsonAry.add(docASTORI);
				// jsonObj.put(rs.getString("ASTORI").trim(), docObj);

				map.put(rs.getString("ASTORI").trim(), docObj);
				listTori.add(rs.getString("ASTORI").trim());
			}
			log.info("map.size=" + map.size());
//			jsonObj.put("begin", this.getsBegin());
//			jsonObj.put("end", this.getsEnd());
//			jsonObj.put("data", jsonAry);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();

			FileUtil.write(file, fileObj.toString(), "UTF-8");

			// log.info("aryyyyyyyyyyyy666=" + obj.get("data").toString());
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
			double sum = 0;
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

				sum += rs.getDouble("ASKGAK");
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
			fillData(jsonobj, sPath + code + "-" + factory + "-" + sBegin + "-" + sEnd + "."
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
				sheetAt.getRow(rowNum).getCell(6).setCellValue(itemObj.getDoubleValue("ASTANK"));
				sheetAt.getRow(rowNum).getCell(7).setCellValue(itemObj.getDoubleValue("ASKGAK"));
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

	public void cleanSendFlag(String path) {
		File jsonfile = new File(path + "record.json");
		if (!jsonfile.exists()) {
			log.info("error!!!transAndSend json文件不存在。");
			return;
		}

		JSONObject fileObj = JSONObject.parseObject(FileUtil.read(jsonfile, "UTF-8"));
		Iterator iter = fileObj.entrySet().iterator();
		while (iter.hasNext()) {
			Map.Entry entry = (Map.Entry) iter.next();
			System.out.println(entry.getKey().toString() + ":" + entry.getValue().toString());
			JSONObject item = (JSONObject) entry.getValue();
			item.put("SEND", 0);
		}
		FileUtil.write(jsonfile, fileObj.toString(), "UTF-8");
		try {
			websckt.sendMessage("FIN|CLNSENDFLAG");
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

		JSONObject fileObj = JSONObject.parseObject(FileUtil.read(jsonfile, "UTF-8"));
//		String sBegin = obj.get("begin").toString();
//		String sEnd = obj.get("end").toString();
//		JSONArray objAry = obj.getJSONArray("data");
		System.out.println("111111111=" + fileObj.size());
		Iterator iter = fileObj.entrySet().iterator();

		try {
			jacobInit();
			int idx = 0;
			int sendflag = 0;
			while (iter.hasNext()) {
				Map.Entry entry = (Map.Entry) iter.next();
				System.out.println(entry.getKey().toString() + ":" + entry.getValue().toString());
				JSONObject item = (JSONObject) entry.getValue();
				JSONArray aryfile = (JSONArray) item.get("FileNames");

				int send = (int) item.get("SEND");
				if (send == 0) {
					// xls转pdf文件
					JSONArray attachfile = new JSONArray();

					for (int i = 0; i < aryfile.size(); i++) {
						String srcFile = path + aryfile.getString(i) + "."
								+ FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE);
						String destFile = path + aryfile.getString(i) + ".pdf";
						File xlsfile = new File(srcFile);
						if (xlsfile.exists()) {
							log.info("存在文件:" + srcFile);
							jacobXls2PDF(srcFile, destFile);
							attachfile.add(destFile);
						}
					}
					// 发邮件
					log.info("ASTORIiiiiiii:" + item.getString("ETPATN") + "," + item.getString("ETPOTO"));
					sendMmczMail(item.getString("ASTORI"), item.getString("ETPATN"), item.getString("ETPOTO"),
							attachfile);
					// sendMmczMail(item.getString("ASTORI"), "xiaojunxie@mektec.com.cn",
					// "911xie@163.com", attachfile);
					item.put("SEND", 1);
					websckt.sendMessage(String.format("SYN|%02d|邮件已发送至供应商[%s](%d/%d)",
							((idx + 1) * 100 / fileObj.size()), item.getString("ASTORI"), (idx + 1), fileObj.size()));
					idx++;
					sendflag = 1;
				}
			}

			jacobUnInit();
			FileUtil.write(jsonfile, fileObj.toString(), "UTF-8");
			String sendmsg = sendflag == 1 ? "FIN|2PDF" : "FIN|SEND1";
			websckt.sendMessage(sendmsg);

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		System.out.println("222222222=" + fileObj.size());

//		int nTotalPdf = 0;
//		for (int i = 0; i < listFiles.length; i++) {
//			String suffix = FilenameUtils.getExtension(listFiles[i].getName()).toLowerCase();
//			File f = listFiles[i];
//			if (f.isFile() && f.getName().indexOf("~$") != 0
//					&& suffix.equalsIgnoreCase(FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE))) {
//				nTotalPdf++;
//			}
//		}
//
//		log.info("transAndSend jsonfilecnt=" + objAry.size() + ",inpathcnt=" + nTotalPdf);
//		if (nTotalPdf != objAry.size()) {
//			log.info("error!!!transAndSend 文件数量对不上。");
//			return;
//		}
//
//		try {
//			jacobInit();
//			for (int i = 0; i < objAry.size(); i++) {
//				JSONObject item = objAry.getJSONObject(i);
//				String baseName = item.getString("ASTORI") + "-" + sBegin + "-" + sEnd;
//				String srcFile = path + baseName + "." + FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE);
//				String destFile = path + baseName + ".pdf";
//				File xlsfile = new File(path + baseName + "." + FilenameUtils.getExtension(BillsDomestic.XLS_TEMPLATE));
//				if (xlsfile.exists()) {
//					log.info("存在文件:" + baseName);
//					jacobXls2PDF(srcFile, destFile);
//					if (item.getString("ETPATN").length() > 0) {
//						log.info("发送邮件:" + item.getString("ASTORI") + "," + item.getString("ETPATN") + ","
//								+ item.getString("ETPOTO"));
//						JSONArray aryfile = new JSONArray();
//						aryfile.add(destFile);
//						sendMmczMail(item.getString("ASTORI"), item.getString("ETPATN"), item.getString("ETPOTO"),
//								aryfile);//
//					}
//					websckt.sendMessage(String.format("SYN|%02d|生成文件[%s](%d/%d)", ((i + 1) * 100 / nTotalPdf),
//							baseName + ".pdf", (i + 1), nTotalPdf));
//				}
//			}
//
//			jacobUnInit();
//			websckt.sendMessage("FIN|2PDF");
//
//		} catch (IOException e) {
//			// TODO Auto-generated catch block
//			e.printStackTrace();
//		}
//		log.info("spare.1111111111=" + (System.currentTimeMillis() - begin));

	}

	public void sendMmczMail(String supplierCode, String to, String cc, JSONArray aryfile) {
		SendMail mail = new SendMail();
		mail.setHost("mmczsmtp.mmcz.mekjpn.ngnet");
		mail.setAuth(false);
		mail.setUsername("");
		mail.setPassword("");
		mail.setFrom(this.sEMailAddr);
		mail.setTo(to);
		String cctarget = cc.trim().length() == 0 ? this.sEMailAddr : cc + "," + this.sEMailAddr;
		mail.setCc(cc);
		log.info("sendMmczMail...cc=" + cctarget);
		mail.setTitle("珠海紫翔对账单" + supplierCode);
		mail.setContent("请查收附件珠海紫翔的对账单，若确认没有问题请及时提供发票,谢谢！");
		mail.setAttachFile(aryfile);
		// mail.setImageFile("d:\\mpechart.png");
		mail.sendMail();
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

		BillsDomestic bills = new BillsDomestic();
		// bills.genToriData("0523", "20220801", "20220815");
		bills.setsEMailAddr("mmcz_monthlystatement@mektec.com.cn");
		JSONArray aryfile = new JSONArray();
		aryfile.add("D:\\eclipse-workspace\\mrrbe\\Reconcile.xltx");
		aryfile.add("D:\\eclipse-workspace\\mrrbe\\Reconcile3.xltx");
		// bills.sendMmczMail("0001", "xiaojunxie@mektec.com.cn",
		// "Xie.XiaoJun@mektec.nokgrp.com", aryfile);
		bills.sendMmczMail("0001", "Xie.XiaoJun@mektec.nokgrp.com", "Xie.XiaoJun@mektec.nokgrp.com", aryfile);
	}

}
