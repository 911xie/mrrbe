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
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
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
public class BillsAbroad implements Runnable {

	private static final Logger log = LoggerFactory.getLogger(Service.class);
	private static final String XLS_TEMPLATE = "Reconcile.xlsx";

	private Service websckt;
	private String director;

	public String getDirector() {
		return director;
	}

	public void setDirector(String director) {
		this.director = director;
	}

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

	private String sCode = "";
	private JSONObject rtnObj;

	BillsAbroad() {
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

			// 测试物料：BETWEEN 20220919 AND 20220920 AND ASTORI='1091'
			// 物料非ALL
			sType = sType.trim().equalsIgnoreCase("ALL") ? "" : " AND MITYPE='" + sType + "' ";
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
					+ ") AA " + "LEFT JOIN YBMETOLA C ON AA.ASTORI=C.ETTORI AND C.ETSSKN='AP' "
					+ "LEFT JOIN YBMTORP D ON AA.ASTORI=D.MVTORI LEFT JOIN UZSKTTP E ON AA.ASTORI=E.TTTORI "
					+ "WHERE AA.SUMASGKIN>0 AND MVADCD='A' " + director + " ORDER BY ASTORI";
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
//			sql = "SELECT A.MVTORI,A.MVTMEK,A.MVADCD,A.MVADDK,A.MVTELN,A.MVFAXN,A.MVCMNO,A.MVBIK1,"
//					+ "B.MTDEC2 MVJYOK,C.MTDEC2 MVTRNS,TRIM(D.MTDEC2)||' '||TRIM(E.MTDEC2) MVTERM,"
//					+ "M.M2BANK,M.M2BANM,F.MVBKNM,A.MVBIK2,F.MVBKNO,M.M2ADDR,M.M2IBAN,M.M2SWIF FROM MMTORIP A "
//					+ "LEFT JOIN M2TORIP M ON A.MVTORI = M.M2TORI "
//					+ "LEFT JOIN MMTABLP B ON A.MVJYOK = B.MTKEYC and B.MTPARM = '0020' "
//					+ "LEFT JOIN MMTABLP C ON A.MVTRNS = C.MTKEYC and C.MTPARM = '0030' "
//					+ "LEFT JOIN MMTABLP D ON A.MVTERM = D.MTKEYC and D.MTPARM = '0040' "
//					+ "LEFT JOIN MMTABLP E ON A.MVRDCP = E.MTKEYC and E.MTPARM = '1140' "
//					+ "LEFT JOIN YBMTORP F ON A.MVTORI = F.MVTORI WHERE A.MVTORI = '" + code + "'";
			sql = "SELECT A.MVTORI, A.MVTMEK, A.MVADCD, A.MVADDK, A.MVTELN, A.MVFAXN,"
					+ "G.MTDEC2 MVJYOK, C.MTDEC2 MVTRNS, H.MTDEC2 MVTERM, F.MVCMNO MVCMNO,F.MVBKNM MVBKNM,F.MVBNK2 M2BANK,"
					+ "F.MVBKNO MVBIK2, F.MVBNKD M2ADDR, F.MVIBAN M2IBAN, F.MVSWIF M2SWIF, "
					+ "case J.MCDAYS when NULLIF(J.MCDAYS,NULL) THEN J.MCDAYS else I.MCDAYS end MCDAYS "
					+ "FROM MMTORIP A LEFT JOIN M2TORIP M ON A.MVTORI = M.M2TORI "
					+ "LEFT JOIN MMTABLP B ON A.MVJYOK = B.MTKEYC and B.MTPARM = '0020' "
					+ "LEFT JOIN MMTABLP C ON A.MVTRNS = C.MTKEYC and C.MTPARM = '0030' "
					+ "LEFT JOIN MMTABLP D ON A.MVTERM = D.MTKEYC and D.MTPARM = '0040' "
					+ "LEFT JOIN MMTABLP E ON A.MVRDCP = E.MTKEYC and E.MTPARM = '1140' "
					+ "LEFT JOIN YBMTORP F ON A.MVTORI = F.MVTORI "
					+ "LEFT JOIN MMTABLP G ON G.MTKEYC = F.MVTERM AND G.MTPARM = 'M280' "
					+ "LEFT JOIN MMTABLP H ON H.MTKEYC = F.MVFOBC AND H.MTPARM = 'M200' "
					+ "LEFT JOIN (SELECT * FROM MMTABLP, MMCKBLP WHERE MMTABLP.MTKEYC=MMCKBLP.MCKEYC) I ON F.MVTERM=I.MTKEYC "
					+ "LEFT JOIN MMCKBLP J ON F.MVTERM=J.MCKEYC " + "WHERE A.MVTORI = '" + code + "'";
			log.info("sql.1=" + sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			// JSONArray jsonarray = new JSONArray();
			JSONObject jsonobj = new JSONObject();
			while (rs.next()) {
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i).toUpperCase();
					jsonobj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}
			}

			int MCDAYS = jsonobj.getInteger("MCDAYS");
			SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyyMMdd");
			SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月");
			Date date = simpleDateFormat.parse(sBegin.substring(0, 6) + "05");
			Calendar calendar = Calendar.getInstance();
			calendar.setTime(date);
			calendar.add(Calendar.DATE, MCDAYS);
			log.info("付款月=" + simpleDateFormat.format(date) + "," + sdf.format(calendar.getTime()));
			jsonobj.put("PAYDATE", sdf.format(calendar.getTime()));

			sql = "SELECT ASKDAY,ASINNO,SUM(ASGKIN) ASGKIN,ASCURR,B.ZKPLNT "
					+ "FROM MTKKTAP A INNER JOIN MZKODRP B ON A.ASPONO=B.ZKPONO and A.ASSEQN=B.ZKSEQN "
					+ "LEFT JOIN MMITEMP C ON A.ASHMCD=C.MIHMCD " + "WHERE ASKDAY BETWEEN " + sBegin + " and " + sEnd
					+ " AND ZKWFG7<>'Y' AND ASTORI='" + code + "'" + sType + sFactory
					+ "GROUP BY ASKDAY,ASINNO,ASCURR,ZKPLNT ORDER BY ASKDAY,ASINNO ";
			log.info("sql.2=" + sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			double Sum1 = 0;
			double Sum2 = 0;
			while (rs.next()) {
				String str = rs.getString("ASKDAY");// "20220802";
				str = str.substring(0, 4) + "/" + str.substring(4, 6) + "/" + str.substring(6, 8);
				jsonobj.put("LastDate", str.trim());
				jsonobj.put("LastUnit", rs.getString("ASCURR").trim());
				jsonobj.put("LastInvNo", rs.getString("ASINNO").trim());

				switch (rs.getInt("ZKPLNT")) {
				case 1:
				case 3:
					// log.info("-----------------01.ASGKIN=" + rs.getFloat("ASGKIN"));
					Sum1 = Sum1 + rs.getFloat("ASGKIN");
					break;
				case 2:
					Sum2 = Sum2 + rs.getFloat("ASGKIN");
					break;
				}
			}
			jsonobj.put("Sum1", Sum1);
			jsonobj.put("Sum2", Sum2);
			log.info("===============Sum1=" + Sum1 + ", Sum2=" + Sum2);

			sql = "SELECT ASORGC,ASTORI, ASKDAY, ASINNO, SUM(ASKGAK) COST, "
					+ "SUM(ASGKIN) INVOICECOST, ASCURR, ASFREC FROM MTKKTAP A "
					+ "INNER JOIN MZKODRP B ON A.ASPONO = B.ZKPONO AND A.ASSEQN = B.ZKSEQN "
					+ "LEFT JOIN MMITEMP C ON A.ASHMCD = C.MIHMCD WHERE ASKDAY BETWEEN " + sBegin + " AND " + sEnd
					+ " AND ASTORI='" + code + "'" + sType + sFactory
					+ "GROUP BY ASORGC,ASTORI,ASKDAY,ASINNO,ASCURR,ASFREC ORDER BY ASORGC,ASTORI,ASKDAY";
			log.info("sql.3=" + sql);
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			JSONArray listData = new JSONArray();
			while (rs.next()) {
				JSONObject objData = new JSONObject();
				objData.put("ASORGC", rs.getString("ASORGC"));
				objData.put("ASTORI", rs.getString("ASTORI"));
				objData.put("ASKDAY", rs.getString("ASKDAY"));
				objData.put("ASINNO", rs.getString("ASINNO").trim());
				objData.put("COST", rs.getString("COST"));
				objData.put("INVOICECOST", rs.getString("INVOICECOST"));
				objData.put("ASCURR", rs.getString("ASCURR"));
				objData.put("ASFREC", rs.getString("ASFREC"));
				listData.add(objData);
			}

			jsonobj.put("listData", listData);
			log.info("ttttttttttt" + code);
			fillData(jsonobj, sPath + code + "-" + sBegin + "-" + sEnd + "."
					+ FilenameUtils.getExtension(BillsAbroad.XLS_TEMPLATE));
		} catch (Exception e) {
			log.info("ddddddddddd" + code + e.getMessage());
			// websckt.sendMessage("ERR|" + e.getMessage());
			rtnObj.put("code", -1);
			rtnObj.put("msg", e.getMessage());
		} finally {
			rs.close();
			stmt.close();
			conn.close();
		}

		return rtnObj;
	}

	public String getStr(String src, int begin, int end) {
		String retStr = "";
		int len = src.length();
		end = end > len ? len : end;
		if (begin >= 0 && begin < src.length() && end > 0 && end <= src.length()) {
			retStr = src.substring(begin, end);
		}
		return retStr;
	}

	public void appendBlankRow(XSSFSheet sheetAt, int num) {
		int beginLen = sheetAt.getLastRowNum() + 1;
		for (int i = 0; i < num - 1; i++) {
			int row = i + beginLen;
			XSSFRow rowObj = sheetAt.createRow(row);
			rowObj.setHeight((short) 450);
			for (int j = 0; j < 12; j++) {
				XSSFCell newcell = rowObj.createCell(j);
			}
		}
	}

	public void appendRow(XSSFSheet sheetAt, int num) {
		int beginLen = sheetAt.getLastRowNum() + 1;
		for (int i = 0; i < num; i++) {
			int row = i + beginLen;
			// 指定合并开始行、合并结束行 合并开始列、合并结束列
			CellRangeAddress range23 = new CellRangeAddress(row, row, 2, 3);
			CellRangeAddress range45 = new CellRangeAddress(row, row, 4, 5);
			CellRangeAddress range67 = new CellRangeAddress(row, row, 6, 7);
			CellRangeAddress range89 = new CellRangeAddress(row, row, 8, 9);
			// 添加要合并地址到表格
			sheetAt.addMergedRegion(range23);
			sheetAt.addMergedRegion(range45);
			sheetAt.addMergedRegion(range67);
			sheetAt.addMergedRegion(range89);
			XSSFRow rowObj = sheetAt.createRow(row);
			rowObj.setHeight((short) 450);
			for (int j = 0; j < 12; j++) {
				XSSFCell newcell = rowObj.createCell(j);
//				if (j == 4 || j == 5) {
//					newcell.setCellType(CellType.NUMERIC);
//				}
			}
		}
	}

	public void editHead(XSSFWorkbook wb, XSSFSheet sheetAt, int beginLen) {

		XSSFFont font = wb.createFont();
		font.setBold(true);

		XSSFCellStyle styRowHead1 = (XSSFCellStyle) wb.createCellStyle();
		styRowHead1.setAlignment(HorizontalAlignment.CENTER);
		styRowHead1.setFont(font);
		styRowHead1.setBorderBottom(BorderStyle.MEDIUM);

		XSSFCellStyle styleCenter = (XSSFCellStyle) wb.createCellStyle();
		styleCenter.setAlignment(HorizontalAlignment.CENTER);
		XSSFCellStyle styleRight = (XSSFCellStyle) wb.createCellStyle();
		styleRight.setAlignment(HorizontalAlignment.RIGHT);

		styleCenter.setFont(font);
		styleRight.setFont(font);

		int rowHead0 = beginLen - 2;
		int rowHead1 = rowHead0 + 1;

		// 指定合并开始行、合并结束行 合并开始列、合并结束列
		CellRangeAddress range01 = new CellRangeAddress(rowHead0, rowHead0, 0, 1);
		CellRangeAddress range26 = new CellRangeAddress(rowHead0, rowHead0, 2, 6);
		CellRangeAddress range710 = new CellRangeAddress(rowHead0, rowHead0, 7, 10);
		// 添加要合并地址到表格
		sheetAt.addMergedRegion(range01);
		sheetAt.addMergedRegion(range26);
		sheetAt.addMergedRegion(range710);

		// 指定合并开始行、合并结束行 合并开始列、合并结束列
		CellRangeAddress range23 = new CellRangeAddress(rowHead1, rowHead1, 2, 3);
		CellRangeAddress range45 = new CellRangeAddress(rowHead1, rowHead1, 4, 5);
		CellRangeAddress range67 = new CellRangeAddress(rowHead1, rowHead1, 6, 7);
		CellRangeAddress range89 = new CellRangeAddress(rowHead1, rowHead1, 8, 9);
		// 添加要合并地址到表格
		sheetAt.addMergedRegion(range23);
		sheetAt.addMergedRegion(range45);
		sheetAt.addMergedRegion(range67);
		sheetAt.addMergedRegion(range89);

		sheetAt.getRow(rowHead0).getCell(0).setCellStyle(styleCenter);
		sheetAt.getRow(rowHead0).getCell(2).setCellStyle(styleCenter);
		sheetAt.getRow(rowHead0).getCell(7).setCellStyle(styleRight);
		sheetAt.getRow(rowHead0).getCell(0).setCellValue("PGM:WB2230R");
		sheetAt.getRow(rowHead0).getCell(2).setCellValue("*** RECIEVING REPORT ***");
		sheetAt.getRow(rowHead0).getCell(7).setCellValue("DATE FROM " + this.getsBegin() + " TO " + this.getsEnd());

		for (int i = 0; i < 11; i++) {
			sheetAt.getRow(rowHead1).getCell(i).setCellStyle(styRowHead1);
		}
		sheetAt.getRow(rowHead1).getCell(0).setCellValue("CODE");
		sheetAt.getRow(rowHead1).getCell(1).setCellValue("VCODE");
		sheetAt.getRow(rowHead1).getCell(2).setCellValue("RECV.DATE");
		sheetAt.getRow(rowHead1).getCell(4).setCellValue("INVOICENO");
		sheetAt.getRow(rowHead1).getCell(6).setCellValue("COST");
		sheetAt.getRow(rowHead1).getCell(8).setCellValue("INVOICECOST");
		sheetAt.getRow(rowHead1).getCell(10).setCellValue("");
		sheetAt.getRow(rowHead1).getCell(11).setCellValue("");
	}

	public void appendMergedRow(XSSFWorkbook wb, XSSFSheet sheetAt, int num) {
		int beginLen = sheetAt.getLastRowNum() + 1;
		XSSFCellStyle styleLine = (XSSFCellStyle) wb.createCellStyle();
		styleLine.setBorderBottom(BorderStyle.MEDIUM);

		for (int i = 0; i < num; i++) {
			int row = i + beginLen;
			// 指定合并开始行、合并结束行 合并开始列、合并结束列
			CellRangeAddress range = new CellRangeAddress(row, row, 0, 11);
			// 添加要合并地址到表格
			sheetAt.addMergedRegion(range);
			XSSFRow rowObj = sheetAt.createRow(row);
			rowObj.setHeight((short) 450);
			for (int j = 0; j < 11; j++) {
				XSSFCell newcell = rowObj.createCell(j);
				newcell.setCellStyle(styleLine);
			}
		}
	}

	public void fillDetail(XSSFWorkbook wb, XSSFSheet sheetAt, JSONArray listData) {
		int nPerPage = 24;
		int nPage = listData.size() / nPerPage;
		log.info("fillDetail.1 page=" + nPage + "," + listData.size());
		nPage = listData.size() % nPerPage > 0 ? nPage + 1 : nPage;
		log.info("fillDetail.2 page=" + nPage);

		// 插入第二页固定空白行好另起一页，经调试8行正好；
		appendBlankRow(sheetAt, 8);

		int beginLen = sheetAt.getLastRowNum();
		int headHeight = 3;
		int tailHeight = 10;
		// appendBlankRow(sheetAt, 10);
		beginLen = beginLen + headHeight;
		int detailHeight = nPerPage;
		XSSFCellStyle style = (XSSFCellStyle) wb.createCellStyle();
		// XSSFCellStyle保留两位小数
		style.setDataFormat(wb.createDataFormat().getFormat("0.00"));

		XSSFCellStyle styleCenter = (XSSFCellStyle) wb.createCellStyle();
		styleCenter.setAlignment(HorizontalAlignment.CENTER);

		try {
			double totalCost = 0;
			double totalInCost = 0;
			for (int n = 0; n < nPage; n++) {
				// 每页页头
				appendBlankRow(sheetAt, headHeight);
				editHead(wb, sheetAt, beginLen);

				// 每页内容
				detailHeight = (n + 1) * nPerPage > listData.size() ? listData.size() - n * nPerPage : detailHeight;
				appendRow(sheetAt, detailHeight);

				for (int i = 0; i < nPerPage; i++) {
					int index = n * nPerPage + i;

					JSONObject itemObj = (JSONObject) listData.get(index);
					if (itemObj != null) {
						int row = i + beginLen;
						sheetAt.getRow(row).getCell(0).setCellStyle(styleCenter);
						sheetAt.getRow(row).getCell(1).setCellStyle(styleCenter);

						sheetAt.getRow(row).getCell(0).setCellValue(itemObj.getString("ASORGC"));
						sheetAt.getRow(row).getCell(1).setCellValue(itemObj.getString("ASTORI"));

						sheetAt.getRow(row).getCell(2).setCellStyle(styleCenter);
						sheetAt.getRow(row).getCell(4).setCellStyle(styleCenter);
						sheetAt.getRow(row).getCell(2).setCellValue(itemObj.getString("ASKDAY"));
						sheetAt.getRow(row).getCell(4).setCellValue(itemObj.getString("ASINNO"));

						sheetAt.getRow(row).getCell(6).setCellStyle(style);
						sheetAt.getRow(row).getCell(8).setCellStyle(style);
						sheetAt.getRow(row).getCell(6).setCellValue(itemObj.getFloat("COST"));
						sheetAt.getRow(row).getCell(8).setCellValue(itemObj.getFloat("INVOICECOST"));
						totalCost += itemObj.getDoubleValue("COST");
						totalInCost += itemObj.getDoubleValue("INVOICECOST");

						sheetAt.getRow(row).getCell(10).setCellValue(itemObj.getString("ASCURR"));
						sheetAt.getRow(row).getCell(11).setCellValue(itemObj.getString("ASFREC"));
						if (index + 1 >= listData.size()) {
							// 统计页
							log.info("fillDetail.2 index=size,break!!!" + index);
							appendMergedRow(wb, sheetAt, 1);
							appendRow(sheetAt, 1);
							// sheetAt.getRow(row + 1).getCell(0).setCellValue("===================");
							sheetAt.getRow(row + 2).getCell(2).setCellValue("TOTAL:");
							sheetAt.getRow(row + 2).getCell(6).setCellStyle(style);
							sheetAt.getRow(row + 2).getCell(8).setCellStyle(style);
							sheetAt.getRow(row + 2).getCell(6).setCellValue(totalCost);
							sheetAt.getRow(row + 2).getCell(8).setCellValue(totalInCost);
							break;
						}
					}
				}
				appendRow(sheetAt, tailHeight);
				beginLen = sheetAt.getLastRowNum() + headHeight;
			}
		} catch (Exception e) {
			log.info("fillDetail.error!" + e.getMessage());
		}
	}

	public void fillData(JSONObject jsonobj, String destFile) throws IOException {
		InputStream in = HikariAS400.class.getClassLoader().getResourceAsStream(BillsAbroad.XLS_TEMPLATE);
		// InputStream in = new FileInputStream(srcFile);
		FileOutputStream out = new FileOutputStream(destFile);
		// 读取整个Excel
		XSSFWorkbook wb = new XSSFWorkbook(in);
		// 获取第一个表单Sheet
		XSSFSheet sheetAt = wb.getSheetAt(0);

		// 默认第一行为标题行，i = 0
		XSSFRow row = sheetAt.getRow(0);
		XSSFCell cell = row.getCell(0);
		cell.setCellValue(jsonobj.getString("MVTMEK"));// ''

		// 单据一
		sheetAt.getRow(0).getCell(0).setCellValue(jsonobj.getString("MVTMEK"));// Master.MVTMEK"MIN AIK PRECISION
																				// INDUSTRIAL CO., LTD."
		String str = jsonobj.getString("MVADDK");// "中国福建省福州市福田区坦洲乡三乡镇一七村台湾街道办桃园观音工业区国瑞路２号";// Master.MVADDK
		sheetAt.getRow(1).getCell(0).setCellValue(getStr(str, 0, 15));
		sheetAt.getRow(2).getCell(0).setCellValue(getStr(str, 15, 30));
		sheetAt.getRow(3).getCell(0).setCellValue(getStr(str, 30, 45));
		sheetAt.getRow(2).getCell(11).setCellValue(jsonobj.getString("LastInvNo"));// Slave.ASINNO
		sheetAt.getRow(3).getCell(11).setCellValue(jsonobj.getString("LastDate"));// Slave.ASKDAY

		sheetAt.getRow(4).getCell(1).setCellValue(jsonobj.getString("MVTELN"));// Master.MVTELN
		sheetAt.getRow(4).getCell(6).setCellValue(jsonobj.getString("MVFAXN"));// Master.MVFAXN

		sheetAt.getRow(9).getCell(2).setCellValue(jsonobj.getString("MVJYOK"));// Master.MVJYOK->YBMTORP.MVTERM
																				// (MMTABLP.MTKEYC =YBMTORP.MVTERM)
																				// 当MTPARM="M280" 取MTDEC2的值
		sheetAt.getRow(10).getCell(2).setCellValue(jsonobj.getString("MVTERM"));// Master.MVTERM->YBMTORP.MVFOBC
																				// (MMTABLP.MTKEYC = YBMTORP.MVFOBC)
																				// 当MTPARM="M200",取MTDEC2的值
		sheetAt.getRow(11).getCell(2).setCellValue(jsonobj.getString("MVCMNO"));// Master.MVCMNO->YBMTORP.MVCMNO
		sheetAt.getRow(12).getCell(2).setCellValue(jsonobj.getString("MVBKNM"));// Master.M2BANM->YBMTORP.MVBKNM
		sheetAt.getRow(13).getCell(2).setCellValue(jsonobj.getString("M2BANK"));// Master.M2BANK->YBMTORP.MVBNK2
		sheetAt.getRow(14).getCell(2).setCellValue(jsonobj.getString("MVBIK2"));// Master.MVBIK2->YBMTORP.MVBKNO
		sheetAt.getRow(15).getCell(2).setCellValue(jsonobj.getString("M2ADDR"));// Master.M2ADDR->YBMTORP.MVBNKD
		sheetAt.getRow(16).getCell(2).setCellValue(jsonobj.getString("M2IBAN"));// Master.M2IBAN->YBMTORP.MVIBAN
		sheetAt.getRow(17).getCell(2).setCellValue(jsonobj.getString("M2SWIF"));// Master.M2SWIF->YBMTORP.MVSWIF

		sheetAt.getRow(20).getCell(8).setCellValue("(" + jsonobj.getString("LastUnit") + ")");// Slave.ASCURR
		sheetAt.getRow(20).getCell(10).setCellValue("(" + jsonobj.getString("LastUnit") + ")");// Slave.ASCURR

		// 数值的不能用字串，否则xls里如果有公式的就不会自动计算了
		sheetAt.getRow(21).getCell(10).setCellValue(jsonobj.getDoubleValue("Sum1"));// Sum1:ZKPLNT==1or3->ASGKIN
		sheetAt.getRow(22).getCell(10).setCellValue(jsonobj.getDoubleValue("Sum2"));// Sum2:ZKPLNT==2->ASGKIN
		sheetAt.getRow(35).getCell(10).setCellValue(jsonobj.getString("PAYDATE"));

		// 单据二
		sheetAt.getRow(38).getCell(11).setCellValue(jsonobj.getString("LastDate"));// Slave.ASKDAY
		sheetAt.getRow(40).getCell(2).setCellValue(jsonobj.getString("MVTMEK"));// Master.MVTMEK
		sheetAt.getRow(42).getCell(2).setCellValue(str);// Master.MVADDK

		sheetAt.getRow(46).getCell(2).setCellValue(jsonobj.getString("MVTRNS"));// Master.MVTRNS
		sheetAt.getRow(48).getCell(2).setCellValue(jsonobj.getString("MVJYOK"));// Master.MVJYOK
		sheetAt.getRow(49).getCell(2).setCellValue(jsonobj.getString("MVTERM"));// Master.MVTERM
		sheetAt.getRow(52).getCell(8).setCellValue("(" + jsonobj.getString("LastUnit") + ")");// Slave.ASCURR
		sheetAt.getRow(52).getCell(10).setCellValue("(" + jsonobj.getString("LastUnit") + ")");// Slave.ASCURR
		sheetAt.getRow(53).getCell(10).setCellValue(jsonobj.getDoubleValue("Sum1"));// Sum1:ZKPLNT==1or3->ASGKIN
		sheetAt.getRow(54).getCell(10).setCellValue(jsonobj.getDoubleValue("Sum2"));// Sum2:ZKPLNT==2->ASGKIN
		sheetAt.getRow(62).getCell(1).setCellValue(jsonobj.getString("MVTMEK"));// Master.MVTMEK

		sheetAt.setForceFormulaRecalculation(true);

		log.info("currlen=" + sheetAt.getRow(3).getHeight() + "," + sheetAt.getRow(36).getHeight() + ","
				+ sheetAt.getRow(67).getHeight() + "," + sheetAt.getLastRowNum());

//		try {
//			JSONArray listData = (JSONArray) jsonobj.get("listData");
//			int beginLen = 95;
//			log.info("item.size()=" + listData.size());
//			for (int i = 0; i < listData.size(); i++) {
//				JSONObject itemObj = (JSONObject) listData.get(i);
//				sheetAt.getRow(i + beginLen).getCell(0).setCellValue(itemObj.getString("ASORGC"));
//				sheetAt.getRow(i + beginLen).getCell(1).setCellValue(itemObj.getString("ASTORI"));
//				sheetAt.getRow(i + beginLen).getCell(2).setCellValue(itemObj.getString("ASKDAY"));
//				sheetAt.getRow(i + beginLen).getCell(3).setCellValue(itemObj.getString("ASINNO"));
//				sheetAt.getRow(i + beginLen).getCell(4).setCellValue(itemObj.getFloat("COST"));
//				sheetAt.getRow(i + beginLen).getCell(5).setCellValue(itemObj.getFloat("INVOICECOST"));
//				sheetAt.getRow(i + beginLen).getCell(6).setCellValue(itemObj.getString("ASCURR"));
//				sheetAt.getRow(i + beginLen).getCell(7).setCellValue(itemObj.getString("ASFREC"));
//				log.info("currlen=" + (i + beginLen));
//			}
//		} catch (Exception e) {
//			log.info("error!" + e.getMessage());
//		}

		JSONArray listData = (JSONArray) jsonobj.get("listData");
		fillDetail(wb, sheetAt, listData);

		wb.write(out);
		out.flush();
		out.close();
		in.close();
	}

	public void xls2Pdf(String xlsFile, String pdfFile) {
//		Workbook wb = new Workbook();
//		// wb.loadFromFile("D:\\eclipse-workspace\\servlet01\\0523-20220601-20220616.xlsx");
//		wb.loadFromFile(xlsFile);
//		Worksheet sheet = wb.getWorksheets().get(0);
//		PageSetup pageSetup = sheet.getPageSetup();
//		// pageSetup.setTopMargin(3);
//		// 如果不向左偏移，右边会越界导致分页
//		pageSetup.setLeftMargin(-2);// \\MCZSVR06\scan\ZP188006
//		// sheet.saveToPdf("D:\\eclipse-workspace\\servlet01\\0523-20220601-20220616.pdf");
//		// sheet.saveToPdf("\\\\MCZSVR06\\scan\\ZP188006\\0523-20220601-20220616.pdf");
//		sheet.saveToPdf(pdfFile);
//		log.info("finish!");
	}

	public void getFileList() {
		File file = new File(sPath);
		File[] listFiles = file.listFiles();
		for (File f : listFiles) {// 将目录内的内容对象化并遍历
			if (f.isFile()) {
				log.info("文件:" + f.getName() + "," + f.getPath());
			}
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
					&& suffix.equalsIgnoreCase(FilenameUtils.getExtension(BillsAbroad.XLS_TEMPLATE))) {
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
				String srcFile = path + baseName + "." + FilenameUtils.getExtension(BillsAbroad.XLS_TEMPLATE);
				String destFile = path + baseName + ".pdf";
				File xlsfile = new File(path + baseName + "." + FilenameUtils.getExtension(BillsAbroad.XLS_TEMPLATE));
				if (xlsfile.exists()) {
					log.info("存在文件:" + baseName);
					jacobXls2PDF(srcFile, destFile);
					if (item.getString("ETPATN").length() > 0) {
						log.info("发送邮件:" + item.getString("ASTORI") + "," + item.getString("ETPATN") + ","
								+ item.getString("ETPOTO"));
						sendMmczMail(item.getString("ASTORI"), item.getString("ETPATN"), item.getString("ETPOTO"),
								destFile);//
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

	public void sendMmczMail(String supplierCode, String to, String cc, String filename) {
		SendMail mail = new SendMail();
		mail.setHost("mmczsmtp.mmcz.mekjpn.ngnet");
		mail.setAuth(false);
		mail.setUsername("");
		mail.setPassword("");
		mail.setFrom(this.sEMailAddr);
		mail.setTo(to);
		String cctarget = cc.trim().length() == 0 ? this.sEMailAddr : cc + "," + this.sEMailAddr;
		mail.setCc(cctarget);
		log.info("sendMmczMail...cc=" + cctarget);
		mail.setTitle("The Statement of Mektec Manufacturing Corp(Zhuhai)LTD " + supplierCode);
		mail.setContent("Statement as attached, please confirm and sign your company seal back. Best regards，");
		mail.setAttachFile(filename);
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
					log.info("genToriDataaaaaaaaaaaaaaaaa=" + listTori.get(i) + "," + i);
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
		BillsAbroad bills = new BillsAbroad();
		// bills.genToriData("0523", "20220801", "20220815");
		bills.setsEMailAddr("mmcz_monthlystatement@mektec.com.cn");
		bills.sendMmczMail("0001", "xiaojunxie@mektec.com.cn", "mengmenggong@mektec.com.cn",
				"D:\\eclipse-workspace\\mrrbe\\Reconcile.xltx");

	}

}
