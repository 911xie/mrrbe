package com.mmcz.Controller;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.LineIterator;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.system.ApplicationHome;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.common.Constants;
import com.common.response.ErrorCode;
import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;
import com.opslab.helper.CharsetHepler;
import com.opslab.helper.FileHelper;
import com.opslab.helper.HikariAS400;
import com.opslab.helper.HikariLocalMsSql;
import com.opslab.helper.HikariMsSql;
import com.opslab.helper.HikariMsSqlNMIS;
import com.opslab.helper.HikariMySql;
import com.opslab.helper.ZipFileUtil;
import com.opslab.util.FileUtil;
import com.websocket.Service;
import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;

@RestController
@RequestMapping("/common")
public class commonController {

	private static final Logger log = LoggerFactory.getLogger(commonController.class);

	// 登录功能实现
	@RequestMapping(value = "/qryUser", method = RequestMethod.POST)
	public JSONObject qryUser(@RequestBody Map<String, String> data) {
		String empcode = data.get("empcode");

		Connection conn = null;
		Statement statement = null;
		ResultSet rs = null;
		try {

			// 创建connection
			conn = HikariMsSql.getDataSource().getConnection();
			statement = conn.createStatement();

			// 执行sql
			String sql = " select b.newdeptid, a.empid, a.empcode, a.empname,"//
					+ "	d0.deptname as deptname0,"// --公司
					+ "	d.deptname as deptname1,"// --本部
					+ "	d2.deptname as deptname2,"// --部
					+ "	d3.deptname as deptname3,"// 课
					+ "	d4.deptname as deptname4 "// 系
					+ " from tps_empinfo a left join torg_department b   on a.deptautoid=b.autoid "
					+ "		left join torg_department d     on a.topdeptautoid=d.autoid "
					+ "		left join tsys_userinfo g       on a.virtualuserid=g.userid "
					+ "		left join torg_department d2    on b.d2=d2.autoid "
					+ "		left join torg_department d3    on b.d3=d3.autoid "
					+ "		left join torg_department d4    on b.d4=d4.autoid "
					+ "     left join torg_department d0    on b.d0=d0.autoid " + " where ( a.isactive=1 "
					+ "         or ( a.isactive=0  and  a.leaveappdate is not null and a.leaveappdate > GETDATE() - 30)) "
					+ "            and d.deptname <>'其他公司人员' and a.empcode='" + empcode + "'";
			PreparedStatement pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE,
					ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONArray jsonarray = new JSONArray();
			JSONObject jsonobj = new JSONObject();
			JSONObject joResult = new JSONObject();

			while (rs.next()) {
				jsonobj.clear();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonobj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}
				jsonarray.add(jsonobj);
			}
			System.out.println("pool查询数据成功 " + jsonarray.toString());
			joResult.put("status", 0);
			joResult.put("message", jsonarray.size() > 0 ? "存在此工号" : "不存在此工号");
			joResult.put("data", jsonarray);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
			return joResult;
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}

	// 对账查询Reconcile
	@RequestMapping(value = "/reconcile", method = RequestMethod.POST)
	public JSONObject reconcile(@RequestBody Map<String, String> data) {
		String empcode = data.get("empcode");

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			String sBegin = "20220801";
			String sEnd = "20220803";

			// 执行sql
			String sql = "SELECT ASTORI, AA.CNT, "
					+ "case C.ETPATN when NULLIF(C.ETPATN,NULL) THEN C.ETPATN else '' end ETPATN,"
					+ "case C.ETPOTO when NULLIF(C.ETPOTO,NULL) THEN C.ETPOTO else '' end ETPOTO "
					+ "FROM (SELECT ASTORI, COUNT(*) CNT "
					+ "FROM MTKKTAP A,MZKODRP B WHERE A.ASPONO=B.ZKPONO AND A.ASSEQN=B.ZKSEQN " + "AND ASKDAY BETWEEN "
					+ sBegin + " AND " + sEnd + " AND ZKWFG7<>'Y' " + "GROUP BY ASTORI ORDER BY ASTORI) AA "
					+ "LEFT JOIN YBMETOLA C ON AA.ASTORI=C.ETTORI AND C.ETSSKN='AP'";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();

			System.out.println("sql=" + sql);
			rs = pstmt.executeQuery();
			while (rs.next()) {
				JSONObject docObj = new JSONObject();
				docObj.put("ETPATN", rs.getString("ETPATN").trim());
				docObj.put("ETPOTO", rs.getString("ETPOTO").trim());
				jsonObj.put(rs.getString("ASTORI").trim(), docObj);
			}
			System.out.println(jsonObj.toString());

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}

	// 周报
	@RequestMapping(value = "/weekreport", method = RequestMethod.POST)
	public ResultBody weekReport(@RequestBody Map<String, String> data) {
		String empcode = data.get("empcode");

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {

			// 创建connection
			conn = HikariMySql.getDataSource().getConnection();
			stmt = conn.createStatement();

			String sBegin = "20220801";
			String sEnd = "20220803";

			String sql = "SELECT GetRangeByWeek(A.DW),A.DW, A.Vendor, A.Project, A.Total Qty,"
					+ "       case when B.nopass is null then 0 else B.nopass end MPEQty,"
					+ "       case when B.nopass is null then '0%' else concat(round(B.nopass*100/total,4),'%') end PPM "
					+ " FROM" + " ("
					+ "	select DATE_FORMAT(CREATION_DATE,'%x-%v') DW,Vendor,Project,count(PcsBarcode) total"
					+ "	from MPEDetailData where CREATION_DATE>='2023-01-01' and CREATION_DATE<'2023-02-01'"
					+ "	AND MPN IN ('NT3357-30','NT3358-37','NT3359-34','NT3360-35','NT3361-32','NT3362-30','NT3634-58')"
					+ "	AND MPEITEM IS NOT NULL GROUP BY DW,Vendor,Project" + " ) A " + " LEFT JOIN " + " ("
					+ "	select DATE_FORMAT(CREATION_DATE,'%x-%v') DW,count(PcsBarcode) as nopass"
					+ "	from MPEDetailData where CREATION_DATE>='2023-01-01' and CREATION_DATE<'2023-02-01' AND ResultCode=152"
					+ "	AND MPN IN ('NT3357-30','NT3358-37','NT3359-34','NT3360-35','NT3361-32','NT3362-30','NT3634-58')"
					+ "	AND MPEITEM IS NOT NULL GROUP BY DW" + " ) B" + " ON A.DW=B.DW order by DW desc";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();
			JSONArray jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				JSONObject docObj = new JSONObject();
				docObj.put("DW", rs.getString("DW").trim());
				docObj.put("Project", rs.getString("Project").trim());
				docObj.put("Qty", rs.getString("Qty").trim());
				docObj.put("MPEQty", rs.getString("MPEQty").trim());
				docObj.put("PPM", rs.getString("PPM").trim());
				// jsonObj.put("WeekReport", docObj);
				jsonAry.add(docObj);
			}
			jsonObj.put("WeekReport", jsonAry);
			System.out.println(jsonAry.toString());

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}

	// 日报
	@RequestMapping(value = "dayreport", method = RequestMethod.POST)
	public ResultBody dayReport(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariMySql.getDataSource().getConnection();
			stmt = conn.createStatement();

			String sql = "SELECT AA.DD,AA.flexname," + "	case when BB.TOTAL is null then 0 else BB.TOTAL end Total,"
					+ "	case when CC.MPEQty is null then 0 else CC.MPEQty end MPEQty,"
					+ "    case when BB.TOTAL is null then 'No Production' when CC.MPEQty is null then 0 else CC.MPEQty end 'MPE Qty' "
					+ "FROM " + "("
					+ "    SELECT DATE_FORMAT(date, '%Y%m%d') DD,flexname,id FROM MPEFlexName A,MPEDate B"
					+ "    WHERE B.date>='2023-01-01' AND B.date<'2023-01-18'" + ") AA " + "LEFT JOIN ("
					+ "    select DATE_FORMAT(CREATION_DATE, '%Y%m%d') Date,Project,count(PcsBarcode) TOTAL "
					+ "    from MPEDetailData where CREATION_DATE >= '2023-01-01' and CREATION_DATE<'2023-01-18' "
					+ "    AND MPN IN ('NT3357-30','NT3358-37','NT3359-34','NT3360-35','NT3361-32','NT3362-30','NT3634-58') AND MPEITEM IS NOT NULL "
					+ "    group by Date,Vendor,Project " + ") BB ON AA.DD=BB.Date AND AA.flexname=BB.Project "
					+ "LEFT JOIN ("
					+ "    select DATE_FORMAT(CREATION_DATE, '%Y%m%d') Date,Project,count(PcsBarcode) MPEQty "
					+ "    from MPEDetailData where ResultCode=152 and CREATION_DATE >= '2023-01-01' and CREATION_DATE<'2023-01-18' "
					+ "    AND MPN IN ('NT3357-30','NT3358-37','NT3359-34','NT3360-35','NT3361-32','NT3362-30','NT3634-58') AND MPEITEM IS NOT NULL "
					+ "    group by Date,Vendor,Project " + ") CC ON AA.DD=CC.Date AND AA.flexname=CC.Project "
					+ "ORDER BY DD,AA.id ";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();
			JSONArray jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				JSONObject docObj = new JSONObject();
				docObj.put("DD", rs.getString("DD").trim());
				docObj.put("flexname", rs.getString("flexname").trim());
				docObj.put("TOTAL", rs.getString("TOTAL").trim());
				docObj.put("MPEQty", rs.getString("MPEQty").trim());
				jsonAry.add(docObj);
			}
			// jsonObj.put("DayReport", jsonAry);
			// result.setData(jsonObj);
			result.setData(jsonAry);
			System.out.println(result);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	// AS400-出入库
	@RequestMapping(value = "inventory", method = RequestMethod.POST)
	public ResultBody inventory(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		String begin = data.get("begin");
		String end = data.get("end");
		String stat = data.get("stat");

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// 出入库数量、金额
			String sql = "";
			if (stat.equalsIgnoreCase("in"))
				// 入库
				sql = "SELECT ANNKDT Date,SUM(ANNQTY) Goodparts,SUM(ANFQTY) Defective,SUM(ANNQTY*ANTAN1) Amount FROM MTNKTAP "
						+ "WHERE ANNKDT>=" + begin + " AND ANNKDT<=" + end + " GROUP BY ANNKDT ORDER BY ANNKDT";
			else if (stat.equalsIgnoreCase("out"))
				// 出库
				sql = "SELECT AOSKDT Date,SUM(AOJRYO) Goodparts,SUM(AOFQTY) Defective,SUM(AOKING) Amount FROM MTSKTAP "
						+ "WHERE AOSKDT>=" + begin + " AND AOSKDT<=" + end + " GROUP BY AOSKDT ORDER BY AOSKDT";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();
			JSONArray jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				jsonObj = new JSONObject();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
			}
			JSONObject jsonRtn = new JSONObject();
			jsonRtn.put("barData", jsonAry);

			// 两厂出入库数量
			sql = "";
			if (stat.equalsIgnoreCase("in"))
				// 入库
				sql = "SELECT ANORGC FACTORY,ANNKDT DATE,SUM(ANFQTY) QTY FROM MTNKTAP " + "WHERE ANNKDT>=" + begin
						+ " AND ANNKDT<=" + end + " GROUP BY ANORGC,ANNKDT ORDER BY ANORGC,ANNKDT";
			else if (stat.equalsIgnoreCase("out"))
				// 出库
				sql = "SELECT AOORGC FACTORY,AOSKDT DATE,SUM(AOFQTY) QTY FROM MTSKTAP " + "WHERE AOSKDT>=" + begin
						+ " AND AOSKDT<=" + end + " GROUP BY AOORGC,AOSKDT ORDER BY AOORGC,AOSKDT";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			rsmd = rs.getMetaData();
			JSONObject jsonObjZ01 = new JSONObject();
			JSONObject jsonObjZ03 = new JSONObject();
			JSONArray jsonAryZ01 = new JSONArray();
			JSONArray jsonAryZ03 = new JSONArray();
			int z01Total = 0;
			int z03Total = 0;

			System.out.println("sql=" + sql);
			while (rs.next()) {

				if (rs.getString("FACTORY").trim().equalsIgnoreCase("Z01")) {
					jsonObjZ01 = new JSONObject();
					for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
						String colName = rsmd.getColumnName(i);
						jsonObjZ01.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
					}
					jsonAryZ01.add(jsonObjZ01);
					z01Total += rs.getInt("QTY");
				}
				if (rs.getString("FACTORY").trim().equalsIgnoreCase("Z03")) {

					jsonObjZ03 = new JSONObject();
					for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
						String colName = rsmd.getColumnName(i);
						jsonObjZ03.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
					}

					jsonAryZ03.add(jsonObjZ03);
					z03Total += rs.getInt("QTY");
				}

			}
			jsonObj = new JSONObject();
			jsonObj.put("Z01", jsonAryZ01);
			jsonObj.put("Z03", jsonAryZ03);
			jsonObj.put("TOTAL01", z01Total);
			jsonObj.put("TOTAL03", z03Total);
			jsonRtn.put("factoryData", jsonObj);

			// 一段时间内总数
			if (stat.equalsIgnoreCase("in"))
				// 入库
				sql = "SELECT SUM(ANNQTY) Goodparts,SUM(ANFQTY) Defective FROM MTNKTAP WHERE ANNKDT>=" + begin
						+ " AND ANNKDT<=" + end;
			else if (stat.equalsIgnoreCase("out"))
				// 出库
				sql = "SELECT SUM(AOJRYO) Goodparts,SUM(AOFQTY) Defective FROM MTSKTAP WHERE AOSKDT>=" + begin
						+ " AND AOSKDT<=" + end;

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			rsmd = rs.getMetaData();
			jsonObj = new JSONObject();
			jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				jsonObj = new JSONObject();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
			}
			jsonRtn.put("pieData", jsonAry);

			// 今日出入库金额
			sql = "SELECT SUM(ANNQTY) Goodparts, SUM(ANNQTY*ANTAN1) Amount FROM MTNKTAP WHERE ANNKDT>=INTEGER(to_char(current_date,'YYYYMMDD')) "
					+ " union "
					+ "SELECT SUM(AOJRYO) Goodparts,SUM(AOKING) Amount FROM MTSKTAP WHERE AOSKDT>=INTEGER(to_char(current_date,'YYYYMMDD'))";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			rsmd = rs.getMetaData();
			jsonObj = new JSONObject();
			jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				jsonObj = new JSONObject();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
			}
			jsonRtn.put("todayData", jsonAry);

			// 客户别数据
			sql = "SELECT B.CMCSRN,A.AOTKCD,trim(B.CMCSRN)||'('||A.AOTKCD||')' USERNAME,B.CMCSNM,AMOUNT FROM ("
					+ " SELECT AOTKCD, SUM(AOKING) AMOUNT FROM MTSKTAP WHERE AOSKDT>=" + begin + " AND AOSKDT<=" + end
					+ " GROUP BY AOTKCD) A LEFT JOIN YOMTOKP B ON A.AOTKCD=B.CMTOKU ORDER BY AMOUNT DESC";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			rsmd = rs.getMetaData();
			jsonObj = new JSONObject();
			jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				jsonObj = new JSONObject();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
			}
			jsonRtn.put("userData", jsonAry);

			result.setData(jsonRtn);
			System.out.println(result);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

////////////////////////////////////////////////////////////////////////////////////////////////////////////////
	@Resource
	ServerListCfg properties;

	public HikariDataSource getMatDataSource(String ip) {
		HikariDataSource dataSource = null;
		HikariConfig hikariConfig = new HikariConfig();
		hikariConfig.setJdbcUrl(String.format("jdbc:sqlserver://%s:1433;DatabaseName=PIMD", ip));
		hikariConfig.setDriverClassName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		hikariConfig.setUsername("pim");
		hikariConfig.setPassword("pimpass");
		hikariConfig.addDataSourceProperty("cachePrepStmts", "true");
		hikariConfig.addDataSourceProperty("prepStmtCacheSize", "250");
		hikariConfig.addDataSourceProperty("prepStmtCacheSqlLimit", "2048");

		dataSource = new HikariDataSource(hikariConfig);
		return dataSource;
	}

	public HikariDataSource getMesDataSource(String ip) {
		HikariDataSource dataSource = null;
		HikariConfig hikariConfig = new HikariConfig();
		hikariConfig.setJdbcUrl(String.format("jdbc:sqlserver://%s:1433;DatabaseName=MEKDW", ip));
		hikariConfig.setDriverClassName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		hikariConfig.setUsername("pim");
		hikariConfig.setPassword("pimpass");
		hikariConfig.addDataSourceProperty("cachePrepStmts", "true");
		hikariConfig.addDataSourceProperty("prepStmtCacheSize", "250");
		hikariConfig.addDataSourceProperty("prepStmtCacheSqlLimit", "2048");

		dataSource = new HikariDataSource(hikariConfig);
		return dataSource;
	}

	@RequestMapping(value = "testHikariDS", method = RequestMethod.POST)
	public ResultBody testHikariDS(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
//		HikariDataSource dataSource = null;
//		HikariConfig hikariConfig = new HikariConfig();
//		hikariConfig.setJdbcUrl("jdbc:sqlserver://localhost:1433;DatabaseName=DockBar");
//		hikariConfig.setDriverClassName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
//		hikariConfig.setUsername("sa");
//		hikariConfig.setPassword("barcode@2022");
//		hikariConfig.addDataSourceProperty("cachePrepStmts", "true");
//		hikariConfig.addDataSourceProperty("prepStmtCacheSize", "250");
//		hikariConfig.addDataSourceProperty("prepStmtCacheSqlLimit", "2048");

		// dataSource = new HikariDataSource(hikariConfig);

		try {
			Connection conn = HikariLocalMsSql.getDataSource().getConnection();
			Statement stmt = null;
			PreparedStatement pstmt = null;
			ResultSet rs = null;

			String sql = "SELECT * FROM SKOITO";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
			// dataSource.close();

		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return result;
	}

//	@RequestMapping(value = "list2did", method = RequestMethod.POST)
//	public ResultBody list2did(@RequestBody Map<String, String> data) {
//		ResultBody result = ResultBodyUtil.success();
//		Connection conn = null;
//		Statement stmt = null;
//		PreparedStatement pstmt = null;
//		ResultSet rs = null;
//		ResultSetMetaData rsmd;
//		try {
//			// 创建connection
//			conn = HikariAS400.getDataSource().getConnection();
//			stmt = conn.createStatement();
//			stmt.execute("set current schema MSRFLIB");
//
//			// String sql = "SELECT * FROM MZ2DIDP ORDER BY ID DESC fetch first 9 rows
//			// only";
//			String sql = "SELECT * FROM MZ2DIDP ORDER BY ID DESC";
//			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
//			rs = pstmt.executeQuery();
//			JSONObject main = genJsonArybyRS(rs);
//			JSONArray jsonAry = (JSONArray) main.get("datalist");
//
//			result.setData(jsonAry);
//			rs.close();
//			pstmt.close();
//			// 关闭connection
//			conn.close();
//		} catch (SQLException e) {
//			e.printStackTrace();
//		}
//		return result;
//	}

	@RequestMapping(value = "list2did", method = RequestMethod.POST)
	public ResultBody list2did(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		int nPage = Integer.parseInt(data.get("nPage") == null ? "0" : data.get("nPage"));
		int nPerPage = Integer.parseInt(data.get("nPerPage") == null ? "3" : data.get("nPerPage"));

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			String sql = "SELECT COUNT(*) total FROM MZ2DIDP";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("total");
				break;
			}

			sql = "SELECT * FROM MZ2DIDP ORDER BY ID DESC LIMIT " + (nPage * nPerPage) + "," + nPerPage;
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();

			System.out.println(sql);

			JSONObject main = genJsonArybyRS(rs);
			JSONArray jsonAry = (JSONArray) main.get("datalist");

			JSONObject rtnData = new JSONObject();
			rtnData.put("aryData", jsonAry);
			rtnData.put("nTotal", nTotal);

			result.setData(rtnData);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "add2did", method = RequestMethod.POST)
	public ResultBody add2did(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;

		String IDHMCD = data.get("IDHMCD");
		String IDTJYN = data.get("IDTJYN");
		String IDEKOT = data.get("IDEKOT");
		String IDRUID = data.get("IDRUID");
		String IDETNM = data.get("IDETNM");
		System.out.println("add.........." + data.toString());

		try {
			// 1.创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// 2.预编译sql语句
			// String sql = "INSERT INTO MZ2DIDP(ID,IDHMCD,IDTJYN,IDEKOT,IDRUID,IDETNM)
			// VALUES(to_char(current_timestamp ,'YYYYMMDDHHmmssSSS'),?,?,?,?,?)";
			String sql = "INSERT INTO MZ2DIDP(IDHMCD,IDTJYN,IDEKOT,IDRUID,IDETNM,CREATETIME) VALUES(?,?,?,?,?,NOW())";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			// 3.填充
			pstmt.setString(1, IDHMCD);
			pstmt.setString(2, IDTJYN);
			pstmt.setString(3, IDEKOT);
			pstmt.setString(4, IDRUID);
			pstmt.setString(5, IDETNM);
			// 4.执行操作
			pstmt.execute();

			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "del2did", method = RequestMethod.POST)
	public ResultBody del2did(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;
		System.out.println("del2did.........." + data.toString());

		int ID = Integer.parseInt(data.get("id"));

		try {
			// 1.创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// 2.预编译sql语句
			String sql = "DELETE FROM MZ2DIDP WHERE ID=?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			// 3.填充
			// pstmt.setString(1, id);
			pstmt.setInt(1, ID);
			// 4.执行操作
			pstmt.execute();

			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "get2didByID", method = RequestMethod.POST)
	public ResultBody get2didByID(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;

		System.out.println("get2didByID.........." + data.toString());

		int ID = Integer.parseInt(data.get("id"));

		try {
			// 1.创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// 2.预编译sql语句
			String sql = "SELECT * FROM MZ2DIDP WHERE ID=?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			// 3.填充
			pstmt.setInt(1, ID);

			// 4.执行操作
			rs = pstmt.executeQuery();

			JSONObject jsonObj = new JSONObject();
			rsmd = rs.getMetaData();
			while (rs.next()) {

				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}
				break;
			}

			result.setData(jsonObj);
			System.out.println("get2didByID.........." + result.toString());

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "upd2did", method = RequestMethod.POST)
	public ResultBody upd2did(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;

		// String ID = data.get("ID");
		int ID = Integer.parseInt(data.get("ID"));
		String IDHMCD = data.get("IDHMCD");
		String IDTJYN = data.get("IDTJYN");
		String IDEKOT = data.get("IDEKOT");
		String IDRUID = data.get("IDRUID");
		String IDETNM = data.get("IDETNM");
		System.out.println("add.........." + data.toString());

		try {
			// 1.创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// 2.预编译sql语句
			String sql = "UPDATE MZ2DIDP SET IDHMCD=?,IDTJYN=?,IDEKOT=?,IDRUID=?,IDETNM=? WHERE ID=?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			// 3.填充
			pstmt.setString(1, IDHMCD);
			pstmt.setString(2, IDTJYN);
			pstmt.setString(3, IDEKOT);
			pstmt.setString(4, IDRUID);
			pstmt.setString(5, IDETNM);
			pstmt.setInt(6, ID);
			// 4.执行操作
			pstmt.execute();

			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "likeQuery", method = RequestMethod.POST)
	public ResultBody likeQuery(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;

		String text = "%" + data.get("text") + "%";
		System.out.println("add.........." + text);

		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			String sql = "SELECT * FROM MZ2DIDP WHERE ID LIKE ? OR IDHMCD LIKE ? OR IDTJYN LIKE ? "
					+ "OR IDEKOT LIKE ? OR IDETNM LIKE ? OR IDRUID  LIKE ?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setString(1, text);
			pstmt.setString(2, text);
			pstmt.setString(3, text);
			pstmt.setString(4, text);
			pstmt.setString(5, text);
			pstmt.setString(6, text);
			rs = pstmt.executeQuery();
			JSONObject main = genJsonArybyRS(rs);
			JSONArray jsonAry = (JSONArray) main.get("datalist");

			result.setData(jsonAry);
			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	public JSONObject genJsonArybyRS(ResultSet rs) {
		ResultSetMetaData rsmd;
		JSONObject jsonRtn = new JSONObject();
		JSONObject codeMap = new JSONObject();
		JSONArray jsonAry = new JSONArray();
		try {
			rsmd = rs.getMetaData();
			while (rs.next()) {
				JSONObject jsonObj = new JSONObject();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
				codeMap.put(rs.getString("IDEKOT"), rs.getString("IDRUID"));
			}
			jsonRtn.put("datalist", jsonAry);
			jsonRtn.put("codemap", codeMap);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// System.out.println(jsonAry.toString());
		return jsonRtn;
	}

	/*
	 * 
	 */
	public JSONObject fillMatData(JSONArray jsonAllLot) {
		JSONArray rtnAry = new JSONArray();
		JSONObject rtnAryObj = new JSONObject();
		List<String> servers = properties.getServers();
		System.out.println("List<String> names=" + servers.toString());

		try {
			for (int i = 0; i < servers.size(); i++) {
				String serverIP = servers.get(i).toString();
				Connection conn = null;
				Statement stmt = null;
				PreparedStatement pstmt = null;
				ResultSet rs = null;
				HikariDataSource hikariDS = getMatDataSource(serverIP);
				conn = hikariDS.getConnection();

				for (int j = 0; j < jsonAllLot.size(); j++) {
					JSONObject obj = (JSONObject) jsonAllLot.get(j);
					String lotno = obj.get("LOTNO").toString();
					String routeid = obj.get("MATCode").toString();
					int matCount = obj.getIntValue("MATCnt");

					if (matCount > 0)
						continue;

					String sql = "SELECT COUNT(*) CNT FROM PcsResults where LotNo in('" + lotno + "')and RouteId="
							+ routeid;
					if (serverIP.equalsIgnoreCase("192.168.138.245"))
						sql = "SELECT COUNT(*) CNT FROM PcsResults where PcsBarcode in "
								+ "(select PcsBarcode from PcsStatus where LotNo='" + lotno
								+ "' and PcsBarcode !='' and PcsBarcode !='NULL') " + "and RouteId=" + routeid;

					pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					rs = pstmt.executeQuery();
					System.out.println("jsonAllLot.rs=" + serverIP + "," + rs);

					while (rs.next()) {
						System.out.println("result=========LotNo=" + rs.getInt("CNT"));
						obj.put("MATCnt", rs.getInt("CNT"));
					}
				}

				System.out.println("jsonAllLot.sizeeeeeeeeee=" + serverIP + "," + jsonAllLot.size() + "," + rs);
				if (rs != null)
					rs.close();
				else
					System.out.println("serveripipipipipipipip=" + serverIP);
				if (pstmt != null)
					pstmt.close();
				conn.close();
				// 释放池连接
				hikariDS.close();
			}
			System.out.println("mat-array-results=" + jsonAllLot.toString());

			String orgKey = "";
			JSONArray rtnAryObjAry = new JSONArray();
			for (int i = 0; i < jsonAllLot.size(); i++) {
				JSONObject obj = (JSONObject) jsonAllLot.get(i);
				if (!orgKey.equalsIgnoreCase(obj.get("KEY").toString())) {
					if (i > 0) {
						rtnAryObj = new JSONObject();
						rtnAryObj.put(orgKey, rtnAryObjAry);
					}
					rtnAryObjAry = new JSONArray();
					rtnAryObjAry.add(obj);
				} else {
					rtnAryObjAry.add(obj);
					if (i == jsonAllLot.size() - 1) {
						rtnAryObj.put(obj.get("KEY").toString(), rtnAryObjAry);
					}
				}

				orgKey = obj.get("KEY").toString();
			}
			System.out.println("mat-array-results666666=" + rtnAryObj.toString());

		} catch (SQLException e) {
			e.printStackTrace();
		}

		return rtnAryObj;
	}

//	// AS400-MAT-MES
//	@RequestMapping(value = "chart2did", method = RequestMethod.POST)
//	public ResultBody chart2did(@RequestBody Map<String, String> data) {
//		ResultBody result = ResultBodyUtil.success();
//		int nPage = Integer.parseInt(data.get("nPage") == null ? "0" : data.get("nPage"));
//		int nPerPage = Integer.parseInt(data.get("nPerPage") == null ? "3" : data.get("nPerPage"));
//		int nPStatus = Integer.parseInt(data.get("nPStatus") == null ? "2" : data.get("nPStatus"));
//		int nLot = Integer.parseInt(data.get("nLot") == null ? "9" : data.get("nLot"));
//
//		Connection conn = null;
//		Statement stmt = null;
//		PreparedStatement pstmt = null;
//		ResultSet rs = null;
//		ResultSetMetaData rsmd;
//		try {
//			// 创建connection
//			conn = HikariAS400.getDataSource().getConnection();
//			stmt = conn.createStatement();
//			stmt.execute("set current schema MSRFLIB");
//
//			String sql = "SELECT * FROM MZ2DIDP";
//			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
//			rs = pstmt.executeQuery();
//
//			// 1.获取设定的品目-as400工程码-mat工程码
//			System.out.println("sql=" + sql);
//			System.out.println("request: nPage=" + nPage + ", nPerPage=" + nPerPage);
//
//			JSONObject main = genJsonArybyRS(rs);
//			JSONArray jsonAry = (JSONArray) main.get("datalist");
//			JSONObject codemap = (JSONObject) main.get("codemap");
//			System.out.println("jsonAry=" + jsonAry.toString());
//			System.out.println("codemap=" + codemap.toString());
//
//			// 2.逐一获取每个品目9个lot数据
//			int nTotal = jsonAry.size();
//			int nStart = nPage * nPerPage;
//			int nEnd = nStart + nPerPage > jsonAry.size() ? jsonAry.size() : nStart + nPerPage;
//			JSONObject jsonMPNObj = new JSONObject();
//			JSONArray jsonAllLot = new JSONArray();
//
//			// 2.1 获取as400数据
//			JSONArray aryData = new JSONArray();
//			for (int i = nStart; i < nEnd; i++) {
//				System.out.println(" " + i + "," + jsonAry.get(i).toString());
//				JSONObject obj = (JSONObject) jsonAry.get(i);
//
//				String mpn = obj.getString("IDHMCD");// 品目名
//				String engcode = obj.getString("IDEKOT");// 工程名
//				String process = obj.getString("IDTJYN");// 工序
//
//				sql = "SELECT ZHETDT,ZHSZNO||ZHLTNO LOTNO,PWHMCD,PWEKOT,PWTJYN,PWKQTY FROM MZSODRLB A "
//						+ "LEFT JOIN MPWPLNEA B ON A.ZHSZNO=B.PWSZNO AND A.ZHLTNO=B.PWLTNO " + "WHERE ZHHMCD='" + mpn
//						+ "' AND ZHSTC2<='" + nPStatus + "' AND PWEKOT='" + engcode + "' AND " + "PWTJYN='" + process
//						+ "' AND PWKQTY>0 ORDER BY LOTNO";
//				pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
//				rs = pstmt.executeQuery();
//				System.out.println("sql=" + sql);
//
//				int n = 0;
//				JSONArray jsonLotAry = new JSONArray();
//				rsmd = rs.getMetaData();
//				String key = obj.getString("IDHMCD") + "," + obj.getString("IDEKOT") + "(" + obj.getString("IDETNM")
//						+ "),工序" + obj.getString("IDTJYN");
//				while (rs.next() && n < nLot) {
//					JSONObject jsonLotObj = new JSONObject();
//					for (int j = 1; j < rsmd.getColumnCount() + 1; j++) {
//						String colName = rsmd.getColumnName(j);
//						if (colName.equalsIgnoreCase("PWKQTY")) {
//							jsonLotObj.put(colName, rs.getString(colName) == null ? "" : rs.getInt(colName));
//						} else {
//							jsonLotObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
//						}
//					}
//					jsonLotObj.put("KEY", key);
//					jsonLotObj.put("MATCnt", 0);
//					jsonLotObj.put("MATCode", codemap.get(rs.getString("PWEKOT")));
//					jsonLotAry.add(jsonLotObj);
//					jsonAllLot.add(jsonLotObj);
//					n++;
//				}
//				System.out.println("666=" + jsonLotAry.toString());
//				// jsonMPNObj.put(obj.getString("IDHMCD") + obj.getString("IDTJYN"), "");
//				aryData.add(key);
//			}
//
//			JSONObject jsonRtn = new JSONObject();
//			// jsonRtn.put("barData", jsonAry);
//
//			// 2.2 获取mat数据 多个连接没关先暂时屏蔽20230313
//			JSONObject mapData = fillMatData(jsonAllLot);
//
//			JSONObject rtnData = new JSONObject();
//			rtnData.put("mapData", mapData);
//			rtnData.put("aryData", aryData);
//			rtnData.put("nTotal", nTotal);
//
//			result.setData(rtnData);
//
//			rs.close();
//			pstmt.close();
//			// 关闭connection
//			conn.close();
//		} catch (SQLException e) {
//			e.printStackTrace();
//		}
//		return result;
//	}

	// AS400-MAT-MES
	@RequestMapping(value = "chart2did", method = RequestMethod.POST)
	public ResultBody chart2did(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		int nPage = Integer.parseInt(data.get("nPage") == null ? "0" : data.get("nPage"));
		int nPerPage = Integer.parseInt(data.get("nPerPage") == null ? "3" : data.get("nPerPage"));
		int nPStatus = Integer.parseInt(data.get("nPStatus") == null ? "2" : data.get("nPStatus"));
		int nLot = Integer.parseInt(data.get("nLot") == null ? "9" : data.get("nLot"));

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			String sql = "SELECT COUNT(*) total FROM MZ2DIDP";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("total");
				break;
			}

			sql = "SELECT * FROM MZ2DIDP ORDER BY ID DESC LIMIT " + (nPage * nPerPage) + "," + nPerPage;
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();

			// 1.获取设定的品目-as400工程码-mat工程码
			System.out.println("sql=" + sql);
			System.out.println("request: nPage=" + nPage + ", nPerPage=" + nPerPage);

			JSONObject main = genJsonArybyRS(rs);
			JSONArray jsonAry = (JSONArray) main.get("datalist");
			JSONObject codemap = (JSONObject) main.get("codemap");
			System.out.println("jsonAry=" + jsonAry.toString());
			System.out.println("codemap=" + codemap.toString());

			// 2.逐一获取每个品目9个lot数据
			JSONObject jsonMPNObj = new JSONObject();
			JSONArray jsonAllLot = new JSONArray();

			// 2.1 获取as400数据
			JSONArray aryData = new JSONArray();
			for (int i = 0; i < jsonAry.size(); i++) {
				System.out.println(" " + i + "," + jsonAry.get(i).toString());
				JSONObject obj = (JSONObject) jsonAry.get(i);

				String mpn = obj.getString("IDHMCD");// 品目名
				String engcode = obj.getString("IDEKOT");// 工程名
				String process = obj.getString("IDTJYN");// 工序

				sql = "SELECT ZHETDT,ZHSZNO||ZHLTNO LOTNO,PWHMCD,PWEKOT,PWTJYN,PWKQTY FROM MZSODRLB A "
						+ "LEFT JOIN MPWPLNEA B ON A.ZHSZNO=B.PWSZNO AND A.ZHLTNO=B.PWLTNO " + "WHERE ZHHMCD='" + mpn
						+ "' AND ZHSTC2<='" + nPStatus + "' AND PWEKOT='" + engcode + "' AND " + "PWTJYN='" + process
						+ "' AND PWKQTY>0 ORDER BY LOTNO";
				pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
				rs = pstmt.executeQuery();
				System.out.println("sql=" + sql);

				int n = 0;
				JSONArray jsonLotAry = new JSONArray();
				rsmd = rs.getMetaData();
				String key = obj.getString("IDHMCD") + "," + obj.getString("IDEKOT") + "(" + obj.getString("IDETNM")
						+ "),工序" + obj.getString("IDTJYN");
				while (rs.next() && n < nLot) {
					JSONObject jsonLotObj = new JSONObject();
					for (int j = 1; j < rsmd.getColumnCount() + 1; j++) {
						String colName = rsmd.getColumnName(j);
						if (colName.equalsIgnoreCase("PWKQTY")) {
							jsonLotObj.put(colName, rs.getString(colName) == null ? "" : rs.getInt(colName));
						} else {
							jsonLotObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
						}
					}
					jsonLotObj.put("KEY", key);
					jsonLotObj.put("MATCnt", 0);
					jsonLotObj.put("MATCode", codemap.get(rs.getString("PWEKOT")));
					jsonLotAry.add(jsonLotObj);
					jsonAllLot.add(jsonLotObj);
					n++;
				}
				System.out.println("666=" + jsonLotAry.toString());
				// jsonMPNObj.put(obj.getString("IDHMCD") + obj.getString("IDTJYN"), "");
				aryData.add(key);
			}

			JSONObject jsonRtn = new JSONObject();
			// jsonRtn.put("barData", jsonAry);

			// 2.2 获取mat数据 多个连接没关先暂时屏蔽20230313
			JSONObject mapData = fillMatData(jsonAllLot);

			JSONObject rtnData = new JSONObject();
			rtnData.put("mapData", mapData);
			rtnData.put("aryData", aryData);
			rtnData.put("nTotal", nTotal);

			result.setData(rtnData);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	public static int getDaysOfMonth(Date date) {
		Calendar calendar = Calendar.getInstance();
		calendar.setTime(date);
		return calendar.getActualMaximum(Calendar.DAY_OF_MONTH);
	}

	// AS400-order-product-sale
	@RequestMapping(value = "ops", method = RequestMethod.POST)
	public ResultBody ops(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();

		// {"yearmonth":"202303","today":"20230305"}
		SimpleDateFormat sdfYMD = new SimpleDateFormat("yyyyMMdd");
		SimpleDateFormat sdfYM = new SimpleDateFormat("yyyyMM");
		String sYearMonth = data.get("yearmonth") == null ? sdfYM.format(new Date()) : data.get("yearmonth");
		String sToday = data.get("today") == null ? sdfYMD.format(new Date()) : data.get("today");

		System.out.println("in params=" + sYearMonth + "," + sToday);

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		ResultSetMetaData rsmd;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			// String sYearMonth = "202303";
			String sYear = sYearMonth.substring(0, 4);
			String sMonth = sYearMonth.substring(4, 6);
			String sField = "MAMM" + sMonth;

			System.out.println("sYearsMonth=" + sYear + "," + sMonth);

			String sql = "SELECT " + sField + " FROM MMCALTLA WHERE MADATE=" + sYear;
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			String schedule = "";
			while (rs.next()) {
				schedule = rs.getString(sField);
			}

			int nCount = 0;
			int nWorkDay = 0;

			Calendar calendar = Calendar.getInstance();
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMdd");
			Date date = sdf.parse(sToday);
			calendar.setTime(date);
			int dayofmonth = calendar.get(Calendar.DAY_OF_MONTH);
			for (int i = 0; i < schedule.length(); i++) {
				if (schedule.charAt(i) == '-') {
					nCount++;
				}

				if (i < dayofmonth) {
					if (schedule.charAt(i) == '-') {
						nWorkDay++;
					}
				}
			}

			sql = "SELECT M1YYMM, M1KOJO, A.SJ SJ1, A.JH JH1, B.SJ SJ2, B.JH JH2, C.SJ SJ3, C.JH JH3 FROM "
					+ "(SELECT M1YYMM, M1KOJO, SUM(M1SJMN) SJ, SUM(M1JHMN) JH FROM MSJORDP WHERE M1YYMM=? GROUP BY M1YYMM,M1KOJO ORDER BY M1YYMM,M1KOJO) A LEFT JOIN "
					+ "(SELECT M2YYMM, M2KOJO, SUM(M2SJMN) SJ, SUM(M2JHMN) JH FROM MSJPRDP WHERE M2YYMM=? GROUP BY M2YYMM,M2KOJO ORDER BY M2YYMM,M2KOJO) B ON A.M1YYMM=B.M2YYMM AND A.M1KOJO=B.M2KOJO LEFT JOIN "
					+ "(SELECT M3YYMM, M3KOJO, SUM(M3SJMN) SJ, SUM(M3JHMN) JH FROM MSJSHIP WHERE M3YYMM=? GROUP BY M3YYMM,M3KOJO ORDER BY M3YYMM,M3KOJO) C ON A.M1YYMM=C.M3YYMM AND A.M1KOJO=C.M3KOJO "
					+ "UNION"
					+ " SELECT M1YYMM, '03' M1KOJO, A.SJ SJ1, A.JH JH1, B.SJ SJ2, B.JH JH2, C.SJ SJ3, C.JH JH3 FROM "
					+ "(SELECT M1YYMM, SUM(M1SJMN) SJ, SUM(M1JHMN) JH FROM MSJORDP WHERE M1YYMM=? GROUP BY M1YYMM ORDER BY M1YYMM) A LEFT JOIN "
					+ "(SELECT M2YYMM, SUM(M2SJMN) SJ, SUM(M2JHMN) JH FROM MSJPRDP WHERE M2YYMM=? GROUP BY M2YYMM ORDER BY M2YYMM) B ON A.M1YYMM = B.M2YYMM LEFT JOIN "
					+ "(SELECT M3YYMM, SUM(M3SJMN) SJ, SUM(M3JHMN) JH FROM MSJSHIP WHERE M3YYMM=? GROUP BY M3YYMM ORDER BY M3YYMM ) C ON A.M1YYMM = C.M3YYMM "
					+ "ORDER BY M1YYMM, M1KOJO";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setInt(1, 202303);
			pstmt.setInt(2, 202303);
			pstmt.setInt(3, 202303);
			pstmt.setInt(4, 202303);
			pstmt.setInt(5, 202303);
			pstmt.setInt(6, 202303);

			rs = pstmt.executeQuery();

			JSONObject jsonRtn = new JSONObject();
			JSONArray jsonAry = new JSONArray();

			rsmd = rs.getMetaData();
			while (rs.next()) {
				JSONObject jsonObj = new JSONObject();

				jsonObj.put("M1YYMM", rs.getString("M1YYMM"));
				String M1KOJO = rs.getString("M1KOJO");
				if (M1KOJO.equalsIgnoreCase("01"))
					M1KOJO = "南屏";
				else if (M1KOJO.equalsIgnoreCase("02"))
					M1KOJO = "龙山";
				else if (M1KOJO.equalsIgnoreCase("03"))
					M1KOJO = "合计";
				else
					M1KOJO = "unknow";
				jsonObj.put("M1KOJO", M1KOJO);
				jsonObj.put("SJ1", rs.getInt("SJ1"));
				jsonObj.put("SJ2", rs.getInt("SJ2"));
				jsonObj.put("SJ3", rs.getInt("SJ3"));
				jsonObj.put("JH1", rs.getInt("JH1"));
				jsonObj.put("JH2", rs.getInt("JH2"));
				jsonObj.put("JH3", rs.getInt("JH3"));

				// 现计划
				jsonObj.put("XJH1", Math.round(((rs.getInt("JH1") * nWorkDay + 0.0) / nCount)));
				jsonObj.put("XJH2", Math.round(((rs.getInt("JH2") * nWorkDay + 0.0) / nCount)));
				jsonObj.put("XJH3", Math.round(((rs.getInt("JH3") * nWorkDay + 0.0) / nCount)));

				// 达成率
				double dcl1 = rs.getInt("SJ1") * 100.00 / jsonObj.getIntValue("XJH1");
				dcl1 = (new BigDecimal(dcl1)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
				double dcl2 = rs.getInt("SJ2") * 100.00 / jsonObj.getIntValue("XJH2");
				dcl2 = (new BigDecimal(dcl2)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
				double dcl3 = rs.getInt("SJ3") * 100.00 / jsonObj.getIntValue("XJH3");
				dcl3 = (new BigDecimal(dcl3)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();

				// 进步率
				double jbl1 = (rs.getInt("SJ1") * 100.00) / jsonObj.getIntValue("JH1");
				jbl1 = (new BigDecimal(jbl1)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
				double jbl2 = (rs.getInt("SJ2") * 100.00) / jsonObj.getIntValue("JH2");
				jbl2 = (new BigDecimal(jbl2)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
				double jbl3 = (rs.getInt("SJ3") * 100.00) / jsonObj.getIntValue("JH3");
				jbl3 = (new BigDecimal(jbl3)).setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();

				jsonObj.put("DCL1", dcl1);
				jsonObj.put("DCL2", dcl2);
				jsonObj.put("DCL3", dcl3);
				jsonObj.put("JBL1", jbl1);
				jsonObj.put("JBL2", jbl2);
				jsonObj.put("JBL3", jbl3);

				jsonAry.add(jsonObj);
			}
			jsonRtn.put("datalist", jsonAry);
			jsonRtn.put("sYear", sYear);
			jsonRtn.put("sMonth", sMonth);
			jsonRtn.put("schedule", schedule);
			jsonRtn.put("nCount", nCount);
			jsonRtn.put("nWorkDay", nWorkDay);

			result.setData(jsonRtn);

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (ParseException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return result;
	}

	// 改造Arrays。toString
	public static String toHexString(byte[] a) {
		if (a == null)
			return "null";
		int iMax = a.length - 1;
		if (iMax == -1)
			return "[]";

		StringBuilder b = new StringBuilder();
		b.append('[');

		String strHex = "";
		StringBuilder sb = new StringBuilder("");
		for (int i = 0;; i++) {
			strHex = Integer.toHexString(a[i] & 0xFF);
			sb.append((strHex.length() == 1) ? "0" + strHex : strHex);
			b.append(strHex.toUpperCase());
			if (i == iMax)
				return b.append(']').toString();
			b.append(", ");
		}
	}

	// NMIS
	@RequestMapping(value = "nmis", method = RequestMethod.POST)
	public ResultBody nmis(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariMsSqlNMIS.getDataSource().getConnection();
			stmt = conn.createStatement();

//			String sql = "insert into tmcz_section(location_cd,section_cd,section_name) values(?,?,?)";
//			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
//			pstmt.setString(1, "0004");
//			pstmt.setString(2, "0002");
//			String name = new String("王".getBytes("GBK"));
//			pstmt.setString(3, name);
//			int row = pstmt.executeUpdate();
			String sql = "insert into charset(gbkname,utf8name,nvchname) values(?,?,N'試してみます、よろしくお願いいたします')";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			// String namegbk = new String("王".getBytes("GBK"));
			// String namegbk = new String("王".getBytes("UTF-8"), "GBK");
			String namegbk = CharsetHepler.toGBK("王");
			String nameutf8 = "王";
			String test = "测试";
//			System.out.println(Arrays.toString(nameutf8.getBytes("gbk")));// [-78, -30, -54, -44]
//			System.out.println(new String(nameutf8.getBytes("gbk"), "GBK"));// 测试
//			System.out.println(Arrays.toString(namegbk.getBytes("gbk")));
//			System.out.println(Arrays.toString("王".getBytes("gbk")));

			System.out.println(commonController.toHexString(nameutf8.getBytes("gbk")));// [-78, -30, -54, -44]
			System.out.println(commonController.toHexString(nameutf8.getBytes()));
			System.out.println(commonController.toHexString("王".getBytes("gbk")));// [-78, -30, -54, -44]
			System.out.println(commonController.toHexString("王".getBytes()));

			System.out.println(commonController.toHexString(namegbk.getBytes("utf-8")));
			System.out.println(commonController.toHexString(namegbk.getBytes()));
			System.out.println(new String(nameutf8.getBytes("gbk"), "GBK"));// 测试
			System.out.println(new String(nameutf8.getBytes("gbk"), "UTF-8"));// 测试
			System.out.println(commonController.toHexString("王".getBytes("gbk")));

//			// OK
//			pstmt.setBytes(1, nameutf8.getBytes("gbk"));
//			pstmt.setBytes(2, nameutf8.getBytes());
//
//			// OK
//			pstmt.setString(1, nameutf8);
//			pstmt.setString(2, nameutf8);
//
//			// GBK-OK,UTF8-Mistaken
//			pstmt.setBytes(1, nameutf8.getBytes("gbk"));
//			pstmt.setBytes(2, nameutf8.getBytes("gbk"));
//
//			// GBK-UTF8的前两个字节，UTF8-OK
//			pstmt.setBytes(1, nameutf8.getBytes());
//			pstmt.setBytes(2, nameutf8.getBytes());

			pstmt.setString(1, nameutf8);
			pstmt.setString(2, nameutf8);
			// pstmt.setString(3, "ご対応ありがとうございます");

			int row = pstmt.executeUpdate();
			System.out.println("---insert row:" + row);

			sql = "select * from charset where id=(select max(id) id from charset)";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();
			JSONArray jsonAry = new JSONArray();

			System.out.println("sql=" + sql);
			while (rs.next()) {
				jsonObj.clear();
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
				}

				jsonAry.add(jsonObj);
			}

			result.setData(jsonAry);
			System.out.println("return=" + result.toString());
			pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		} catch (UnsupportedEncodingException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "/getbglist", method = RequestMethod.POST)
	private JSONObject getBGList(@RequestBody Map<String, String> data) {
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;

		JSONObject jsonobj = new JSONObject();
		JSONArray jsonAry = new JSONArray();
		try {
			conn = HikariMySql.getDataSource().getConnection();
			String sql = "select * from bgimg";

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();

			while (rs.next()) {
				jsonAry.add(rs.getString("imgname").trim());
			}
			jsonobj.put("list", jsonAry);

			System.out.println(jsonobj.toJSONString());
			rs.close();
			pstmt.close();
			conn.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return jsonobj;
	}

	// QA二维码数据导出
	@RequestMapping(value = "/exportQA", method = RequestMethod.POST)
	private JSONObject exportQA(@RequestBody Map<String, String> data) {
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {

			HikariDataSource hikariDS = getMatDataSource("192.168.138.245");
			conn = hikariDS.getConnection();

			// 执行sql
			String sql = "SELECT COUNT(*) CNT FROM PcsQA";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();

			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("CNT");
				break;
			}
			System.out.println("000000000000:" + nTotal);

			ApplicationHome h = new ApplicationHome(getClass());
			File jarPath = h.getSource();
			String fileName = jarPath.getParentFile().toString() + "\\QAdata.xlsx";
			FileOutputStream fileOutputStream = new FileOutputStream(fileName);
			// XSSFWorkbook workbook = new XSSFWorkbook();
			SXSSFWorkbook workbook = new SXSSFWorkbook();

			int nPerSheet = 1000000;
			int nSheet = nTotal / nPerSheet;
			nSheet = nSheet * nPerSheet < nTotal ? nSheet + 1 : nSheet;

			for (int i = 0; i < nSheet; i++) {
				String sheetName = String.format("QADATA%d", i);
				SXSSFSheet sheet = workbook.createSheet(sheetName);
				SXSSFRow rowTitle = sheet.createRow(0);
				rowTitle.createCell(0).setCellValue("ShtBarcode");
				rowTitle.createCell(1).setCellValue("PcsIndex");
				rowTitle.createCell(2).setCellValue("PcsBarCode");
				rowTitle.createCell(3).setCellValue("LotNo");
				rowTitle.createCell(4).setCellValue("MPN");
				rowTitle.createCell(5).setCellValue("PackingOrderNo");
				rowTitle.createCell(6).setCellValue("CPN1");
				rowTitle.createCell(7).setCellValue("Inspector");
				rowTitle.createCell(8).setCellValue("QualifyNo");
				rowTitle.createCell(9).setCellValue("CPN2");
				rowTitle.createCell(10).setCellValue("DeviceId");
				rowTitle.createCell(11).setCellValue("EnterTime");
				rowTitle.createCell(12).setCellValue("MILotNo");
				rowTitle.createCell(13).setCellValue("PackageNo");
				rowTitle.createCell(14).setCellValue("VacuumVal");
				rowTitle.createCell(15).setCellValue("FPCLOTNO");

				int nPageTotal = (i + 1) * nPerSheet < nTotal ? (i + 1) * nPerSheet : (nTotal - nPerSheet);
				int nTimes = 20000;
				int times = nPageTotal / nTimes;
				times = times * nTimes < nPageTotal ? times + 1 : times;
				int nCurrLine = 1;
				for (int j = 0; j < times; j++) {
					int nBegin = i * nPerSheet + j * nTimes + 1;
					int nEnd = nBegin + nTimes - 1;
					nEnd = nEnd > nTotal ? nTotal : nEnd;

					sql = "SELECT quotename(A.ShtBarcode) AS ShtBarcode,A.PcsIndex,A.PcsBarCode,C.LotNo,C.MPN,B.PackingOrderNo,"
							+ "C.CPN1,C.Inspector,C.QualifyNo,C.CPN2,C.DeviceId,C.EnterTime,C.MILotNo,C.PackageNo,C.VacuumVal,A.LOTNO FPCLOTNO "
							+ "FROM (select * from (select *, ROW_NUMBER() OVER(Order by PcsBarCode) AS RowId from PcsQA) as b where RowId between ? and ?) D "
							+ "LEFT JOIN PcsResults A WITH(NOLOCK) ON D.PcsBarCode=A.PcsBarcode "
							+ "LEFT JOIN PackingOrderDetial B WITH(NOLOCK) ON A.PcsBarcode=B.PcsBarCode "
							+ "LEFT JOIN PackingOrderInfo C WITH(NOLOCK) ON B.PackingOrderNo=C.PackingOrderNo "
							+ "WHERE A.RouteId=350";
					pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					pstmt.setInt(1, nBegin);
					pstmt.setInt(2, nEnd);
					rs = pstmt.executeQuery();
					log.info("1.query finish begin=" + nBegin + ", end=" + nEnd);
					while (rs.next()) {
						SXSSFRow row = sheet.createRow(nCurrLine);
						row.createCell(0).setCellValue(rs.getString("ShtBarcode"));
						row.createCell(1).setCellValue(rs.getInt("PcsIndex"));
						row.createCell(2).setCellValue(rs.getString("PcsBarCode"));
						row.createCell(3).setCellValue(rs.getString("LotNo"));
						row.createCell(4).setCellValue(rs.getString("MPN"));
						row.createCell(5).setCellValue(rs.getInt("PackingOrderNo"));
						row.createCell(6).setCellValue(rs.getString("CPN1"));
						row.createCell(7).setCellValue(rs.getString("Inspector"));
						row.createCell(8).setCellValue(rs.getString("QualifyNo"));
						row.createCell(9).setCellValue(rs.getString("CPN2"));
						row.createCell(10).setCellValue(rs.getInt("DeviceId"));
						row.createCell(11).setCellValue(rs.getString("EnterTime"));
						row.createCell(12).setCellValue(rs.getString("MILotNo"));
						row.createCell(13).setCellValue(rs.getString("PackageNo"));
						row.createCell(14).setCellValue(rs.getString("VacuumVal"));
						row.createCell(15).setCellValue(rs.getString("FPCLOTNO"));
						nCurrLine++;
					}

					log.info("2.sheetNo=" + i + ", begin=" + nBegin + ", end=" + nEnd);
				}
			}
			workbook.write(fileOutputStream);
			fileOutputStream.flush();
			fileOutputStream.close();
			workbook.close();

			rs.close();
			pstmt.close();
			// 关闭connection
			conn.close();
			// 释放池连接
			hikariDS.close();
		} catch (SQLException | IOException e) {
			e.printStackTrace();
		}
		return null;
	}

	// 数据库导出导入程序MES-BI
	@RequestMapping(value = "/mes2bi", method = RequestMethod.POST)
	private JSONObject mes2bi(@RequestBody Map<String, String> data) {
		String begin = data.get("begin");
		String end = data.get("end");

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		LineIterator it = null;
		try {

			HikariDataSource hikariDS = getMesDataSource ("192.168.131.61");
			conn = hikariDS.getConnection();

			// 1.获取总数
			String sql = "select count(*) cnt from MES_CP_Meter where EnterTime>=?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setString(1, begin);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			JSONObject jsonObj = new JSONObject();

			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("cnt");
				break;
			}
			System.out.println("000000000000:" + nTotal);

			ApplicationHome h = new ApplicationHome(getClass());
			File jarPath = h.getSource();
			String fileName = jarPath.getParentFile().toString() + "\\mes2bi.sql";
			FileOutputStream fos = new FileOutputStream(fileName);

			// 2.分页获取，写入sql
			int nPerPage = 10000;
			int nPageCnt = nTotal / nPerPage;
			nPageCnt = nPageCnt * nPerPage < nTotal ? nPageCnt + 1 : nPageCnt;
			for (int i = 0; i < nPageCnt; i++) {
				sql = "select deviceid,EnterTime,FLOOR(DATEPART(HOUR,ENTERTIME)/2) periodindex, ParameterID," +
						"LSL, USL,Value,State from MES_CP_Meter where EnterTime>=? and EnterTime<? " +
						"order by EnterTime offset ? rows fetch next ? rows only";
				pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
				pstmt.setString(1, begin);
				pstmt.setString(2, end);
				pstmt.setInt(3, i * nPerPage);
				pstmt.setInt(4, nPerPage);
				rs = pstmt.executeQuery();
				int lastnum = 0;
				while (rs.next()) {
					String line = "INSERT INTO device_parameter_dashboard_sub(DeviceId,InputDateTime,PeriodIndex," +
							"ParameterId,LimitLower,LimitUpper,ParameterValue,ParameterStatus) VALUES('" +
							rs.getString("deviceid")+"','"+rs.getString("EnterTime")+"',"+
							rs.getString("periodindex")+",'"+rs.getString("ParameterID")+"',"+
							rs.getDouble("LSL")+","+rs.getDouble("USL")+","+
							rs.getDouble("Value")+","+rs.getInt("State")+");\r\n";
					fos.write(line.getBytes("UTF8"));
					lastnum++;
				}
				log.info("1.query finish begin=" + begin + ", page=" + i + ", last=" + lastnum);
			}

			fos.flush();
			fos.close();

			String url = "jdbc:postgresql://192.168.137.25:5432/BIDATA";
			Class.forName("org.postgresql.Driver");
			Connection pgconn = DriverManager.getConnection(url, "zhbi", "zhbi@123");
			File f = new File(fileName);
			it = FileUtils.lineIterator(f, "UTF-8");
			stmt = pgconn.createStatement();
			int nRow = 0;
			long ms = System.currentTimeMillis();
			while (it.hasNext()) {
				String line = it.nextLine();
				stmt.addBatch(line);

				if (nRow % 1000 == 0) {
					stmt.executeBatch();
					stmt.clearBatch();
					log.info("sql="+line);
					log.info("currline="+ nRow + ", time=" + (System.currentTimeMillis() - ms));
					ms = System.currentTimeMillis();
				}

				nRow++;
			}
			stmt.executeBatch();

//			stmt = pgconn.createStatement();
//			stmt.addBatch(sql1);
//			stmt.addBatch(sql2);
//			stmt.executeBatch();
//			stmt.clearBatch();
//			stmt.close();

			rs.close();
			stmt.close();
			pstmt.close();

			// 关闭connection
			conn.close();
			pgconn.close();
			// 释放池连接
			hikariDS.close();
		} catch (SQLException | IOException e) {
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			throw new RuntimeException(e);
		} finally {

			LineIterator.closeQuietly(it);

		}
		return null;
	}

	@RequestMapping(value = "/getOeeCode", method = RequestMethod.POST)
	private ResultBody getOeeCode(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		String devID = data.get("devID");
		String endTime = data.get("endTime");
		String startTime = data.get("startTime");
		log.info("param=" + data.toString());

		try {
			String url = "jdbc:postgresql://192.168.137.145:5412/iot";
			Class.forName("org.postgresql.Driver");
			Connection conn = DriverManager.getConnection(url, "guest", "guest");
			Statement stmt = conn.createStatement();
			PreparedStatement pstmt = null;
			ResultSet rs = null;

			String sql = "select code from CODE A,ITEM B WHERE A.id = B.objectid "
					+ " AND A.value =? AND A.deleted=FALSE order by code";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setString(1, devID);
			rs = pstmt.executeQuery();
			ResultSetMetaData rsmd = rs.getMetaData();
			int index = 0;
			JSONArray jsonAry = new JSONArray();

			rsmd = rs.getMetaData();
			while (rs.next()) {
				JSONObject jsonObj = new JSONObject();
				jsonObj.put("value", rs.getString("code"));
				jsonObj.put("label", rs.getString("code"));
				jsonAry.add(jsonObj);
			}

			result.setData(jsonAry);
			log.info("result=" + result.toString());

			rs.close();
			stmt.close();
			conn.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;
	}

	// 获取OEE数据
	@RequestMapping(value = "/getOeeData", method = RequestMethod.POST)
	private ResultBody getOeeData(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		int nPage = Integer.parseInt(data.get("nPage") == null ? "0" : data.get("nPage"));
		int nPerPage = Integer.parseInt(data.get("nPerPage") == null ? "10" : data.get("nPerPage"));

		String devID = data.get("devID");
		String endTime = data.get("endTime");
		String startTime = data.get("startTime");
		log.info("param=" + data.toString());

		try {
			String url = "jdbc:postgresql://192.168.137.145:5412/iot";
			Class.forName("org.postgresql.Driver");
			Connection conn = DriverManager.getConnection(url, "guest", "guest");
			PreparedStatement pstmt = null;
			ResultSet rs = null;
			String sql = "SELECT COUNT(*) nTotal " + "FROM stateEvent A, state B, statetype C, code D "
					+ "WHERE B.id = A.stateid AND B.typeid = C.id AND B.objectid = D.target "
					+ "AND D.value = ? AND A.starttime >= ? AND A.endtime <= ?  ";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setString(1, devID);
			Timestamp st = Timestamp.valueOf(startTime);
			Timestamp et = Timestamp.valueOf(endTime);
			pstmt.setTimestamp(2, st);
			pstmt.setTimestamp(3, et);
			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("nTotal");
				break;
			}

			JSONArray jsonAry = new JSONArray();
			if (nTotal > 0) {
				int nMaxPage = (nTotal / nPerPage) * nPerPage >= nTotal ? (nTotal / nPerPage) : (nTotal / nPerPage + 1);
				nMaxPage--;
				log.info("nMaxPage = " + nMaxPage + ", req page=" + nPage + ",perPage=" + nPerPage);
				nPage = nPage > nMaxPage ? nMaxPage : nPage;

				sql = "SELECT A.id,D.value, C.categoryid, C.name, A.starttime, A.endtime, A.duration "
						+ "FROM stateEvent A, state B, statetype C, code D "
						+ "WHERE B.id = A.stateid AND B.typeid = C.id AND B.objectid = D.target "
						+ "AND D.value = ? AND A.starttime >= ? AND A.endtime <= ? order by endtime desc "
						+ "OFFSET ? LIMIT ?";
				pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
				pstmt.setString(1, devID);
				pstmt.setTimestamp(2, st);
				pstmt.setTimestamp(3, et);
				pstmt.setInt(4, nPage * nPerPage);
				pstmt.setInt(5, nPerPage);
				rs = pstmt.executeQuery();
				ResultSetMetaData rsmd = rs.getMetaData();
				int index = 0;

				rsmd = rs.getMetaData();
				while (rs.next()) {
					JSONObject jsonObj = new JSONObject();
					for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
						String colName = rsmd.getColumnName(i);
						jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
					}

					jsonAry.add(jsonObj);
				}
			}

			log.info("jsonAry-length = " + jsonAry.size());
			JSONObject rtnData = new JSONObject();
			rtnData.put("aryData", jsonAry);
			rtnData.put("nTotal", nTotal);
			rtnData.put("reqPage", nPage);
			result.setData(rtnData);

			if (rs != null)
				rs.close();
			conn.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return result;

	}

	@RequestMapping(value = "/material", method = RequestMethod.POST)
	public ResultBody material(@RequestBody Map<String, String> data) {
		ResultBody result = ResultBodyUtil.success();
		int nPage = Integer.parseInt(data.get("nPage") == null ? "0" : data.get("nPage"));
		int nPerPage = Integer.parseInt(data.get("nPerPage") == null ? "10" : data.get("nPerPage"));

		String LotNo = data.get("LotNo");
		int nStart = Integer.parseInt(data.get("starttime") == null ? "20230101" : data.get("starttime"));

		if (LotNo.length() != 11){
			result = ResultBodyUtil.fail(ErrorCode.ILLEGAL_PARAMETER, "Lot号不正确");
			return result;
		}

		long ms = System.currentTimeMillis();
		log.info("intime=" + ms);
		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;
		try {
			// 创建connection
			conn = HikariAS400.getDataSource().getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");
			String sql = "SELECT COUNT(*) nTotal FROM MTCMPXP WHERE BXPKNO=? AND BXCPDT>=?";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.setString(1, LotNo);
			pstmt.setInt(2, nStart);
			rs = pstmt.executeQuery();
			int nTotal = 0;
			while (rs.next()) {
				nTotal = rs.getInt("nTotal");
				break;
			}
			//pstmt.close();
			log.info("1.gettotal=" + nTotal+", times="+(System.currentTimeMillis() - ms));

			JSONArray jsonAry = new JSONArray();
			if (nTotal > 0) {
				int nMaxPage = (nTotal / nPerPage) * nPerPage >= nTotal ? (nTotal / nPerPage) : (nTotal / nPerPage + 1);
				nMaxPage--;
				nPage = nPage > nMaxPage ? nMaxPage : nPage;
				// 执行sql
//				sql = "SELECT BXPKNO,BXHMCD,BXTJYN,BXKOTE,C.EBNAME,BXZAIR,B.POPART,BXLOTE," +
//						" D.T1FDNO,E.JPSLOT,BXBASY,BXKQTY,BXTANI,BXCPDT,BXCPTM,BXSAGY, " +
//						" TO_CHAR(to_date(BXCPDT||digits(cast(BXCPTM as decimal(6))),'yyyymmddhh24miss'),'yyyy-mm-dd hh24:mi:ss') AS CREATEDATE " +
//						" FROM " +
//						" (SELECT * FROM MTCMPXP WHERE BXPKNO=?) A" +
//						" LEFT JOIN" +
//						" (SELECT POSZNO||POLTNO LOTNO,POTJYN,POPART,POHMCD,POPACD FROM MPOPLNP" +
//						" WHERE POSZNO=substr(?,1,8) AND POLTNO=substr(?,9,3)) B" +
//						" ON A.BXPKNO=B.LOTNO AND A.BXTJYN=B.POTJYN AND A.BXZAIR=B.POHMCD" +
//						" LEFT JOIN EBTTBGP C ON A.BXKOTE=C.EBCODE" +
//						" LEFT JOIN MTCMP1LA D ON A.BXLOTE=D.T1SZLT" +
//						" LEFT JOIN MMMJSTP E ON A.BXTJYN=E.JPTJYN AND A.BXZAIR=E.JPZAIR" +
//						" WHERE BXCPDT>=? ORDER BY BXCPDT,BXCPTM LIMIT ?,?";

				int preMode = 0;
				if (preMode == 1) {
					sql = "SELECT BXPKNO,BXHMCD,BXTJYN,BXKOTE,C.EBNAME,BXZAIR,B.POPART," +
							" BXLOTE,BXBASY,BXKQTY,BXTANI,BXCPDT,BXCPTM,BXSAGY, " +
							" TO_CHAR(to_date(BXCPDT||digits(cast(BXCPTM as decimal(6))),'yyyymmddhh24miss'),'yyyy-mm-dd hh24:mi:ss') AS CREATEDATE " +
							" FROM " +
							" (SELECT * FROM MTCMPXP WHERE BXPKNO=?) A" +
							" LEFT JOIN" +
							" (SELECT POSZNO||POLTNO LOTNO,POTJYN,POPART,POHMCD,POPACD FROM MPOPLNP" +
							" WHERE POSZNO=substr(?,1,8) AND POLTNO=substr(?,9,3)) B" +
							" ON A.BXPKNO=B.LOTNO AND A.BXTJYN=B.POTJYN AND A.BXZAIR=B.POHMCD" +
							" LEFT JOIN EBTTBGP C ON A.BXKOTE=C.EBCODE" +
							" WHERE BXCPDT>=? ORDER BY BXCPDT,BXCPTM LIMIT ?,?";
					pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
					pstmt.setString(1, LotNo);
					pstmt.setString(2, LotNo);
					pstmt.setString(3, LotNo);
					pstmt.setInt(4, nStart);
					pstmt.setInt(5, nPage * nPerPage);
					pstmt.setInt(6, nPerPage);
				}
				else {
					sql = "SELECT BXPKNO,BXHMCD,BXTJYN,BXKOTE,C.EBNAME,BXZAIR,B.POPART,\n" +
							" BXLOTE,BXBASY,BXKQTY,BXTANI,BXCPDT,BXCPTM,BXSAGY, \n" +
							" TO_CHAR(to_date(BXCPDT||digits(cast(BXCPTM as decimal(6))),'yyyymmddhh24miss'),'yyyy-mm-dd hh24:mi:ss') AS CREATEDATE \n" +
							" FROM \n" +
							" (SELECT * FROM MTCMPXP WHERE BXPKNO='" + LotNo + "') A\n" +
							" LEFT JOIN\n" +
							" (SELECT POSZNO||POLTNO LOTNO,POTJYN,POPART,POHMCD,POPACD FROM MPOPLNP\n" +
							" WHERE POSZNO=substr('" + LotNo + "',1,8) AND POLTNO=substr('" + LotNo + "',9,3)) B\n" +
							" ON A.BXPKNO=B.LOTNO AND A.BXTJYN=B.POTJYN AND A.BXZAIR=B.POHMCD\n" +
							" LEFT JOIN EBTTBGP C ON A.BXKOTE=C.EBCODE\n" +
							" WHERE BXCPDT>=" + nStart + " ORDER BY BXCPDT,BXCPTM LIMIT " + nPage * nPerPage + "," + nPerPage;
					pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
				}

				log.info("sql(mode=" + preMode + ") = "+sql);
				log.info("2.0.nStart = " + nStart + ", nPage = " + nPage +", limit begin=" + nPage * nPerPage + ",perPage=" + nPerPage);
				log.info("2.1.getdata time=" + (System.currentTimeMillis() - ms));
				rs = pstmt.executeQuery();
				log.info("2.2.getdata time=" + (System.currentTimeMillis() - ms));
				ResultSetMetaData rsmd = rs.getMetaData();
				log.info("2.3.getdata time=" + (System.currentTimeMillis() - ms));

				//System.out.println("sql=" + sql);
				while (rs.next()) {
					JSONObject jsonObj = new JSONObject();
					for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
						String colName = rsmd.getColumnName(i);
						jsonObj.put(colName, rs.getString(colName) == null ? "" : rs.getString(colName).trim());
					}

					jsonAry.add(jsonObj);
				}
			}
			log.info("3.getdata time=" + (System.currentTimeMillis() - ms));

			JSONObject rtnData = new JSONObject();
			rtnData.put("aryData", jsonAry);
			rtnData.put("nTotal", nTotal);
			rtnData.put("reqPage", nPage);
			result.setData(rtnData);

			rs.close();
			//pstmt.close();
			// 关闭connection
			conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return result;
	}

	@RequestMapping(value = "import2did", method = RequestMethod.POST)
	// @RequestMapping("/import2did")
	@ResponseBody
	public ResultBody import2did(@RequestPart("single") MultipartFile mf, @RequestPart("data") String dataPath) {
		ResultBody result = ResultBodyUtil.success();
		ApplicationHome h = new ApplicationHome(getClass());
		File jarPath = h.getSource();

		// String fileName = jarPath.getParentFile().toString() + "\\" +
		// mf.getOriginalFilename();
		String fileName = jarPath.getParentFile().toString() + "\\" + UUID.randomUUID();
		String destFileName = fileName;

		System.out.println("fileName:" + fileName);

		Connection conn = null;
		Statement stmt = null;
		PreparedStatement pstmt = null;
		ResultSet rs = null;

		try {
			BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(destFileName));
			out.write(mf.getBytes());
			out.flush();
			out.close();
			log.info("上传文件完成：" + mf.getOriginalFilename() + "===>" + destFileName);

			File f = new File(fileName);
			InputStream in = new FileInputStream(f);
			// 读取整个Excel
			XSSFWorkbook wb = new XSSFWorkbook(in);
			// 获取第一个表单Sheet
			XSSFSheet sheetAt = wb.getSheetAt(0);
			XSSFRow row0 = sheetAt.getRow(0);
			String sql = "INSERT INTO SESSION.TMP2DID(";
			String col = "";
			// 表头
			for (int i = 0; i < row0.getPhysicalNumberOfCells(); i++) {
				XSSFCell cell = row0.getCell(i);
				col += cell.getStringCellValue() + ",";
			}
			col += "CREATETIME";
			// col = col.substring(0, col.length() - 1);
			sql = sql + col + ") VALUES ";

			// 数据
			int nRow = sheetAt.getPhysicalNumberOfRows();
			String sdata = "";
			for (int i = 1; i < nRow; i++) {
				XSSFRow row = sheetAt.getRow(i);
				XSSFCell cell0 = row.getCell(0);
				if (cell0.getStringCellValue().trim().length() > 0) {
					sdata += "(";
					for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
						XSSFCell cell = row.getCell(j);
						cell.setCellType(CellType.STRING);
						sdata += "'" + cell.getStringCellValue() + "',";
					}
					// sdata = sdata.substring(0, sdata.length() - 1);
					sdata += "NOW()),";
				}
			}

			sdata = sdata.substring(0, sdata.length() - 1);
			sql += sdata;

			wb.close();
			in.close();

			System.out.println("cell :" + sql + ",row=" + nRow);

			HikariConfig hikariConfig = new HikariConfig();
			hikariConfig.setJdbcUrl("jdbc:as400://192.168.121.40");
			hikariConfig.setDriverClassName("com.ibm.as400.access.AS400JDBCDriver");
			hikariConfig.setUsername("DENSIO");
			hikariConfig.setPassword("ODBC");
			hikariConfig.addDataSourceProperty("cachePrepStmts", "true");
			hikariConfig.addDataSourceProperty("prepStmtCacheSize", "250");
			hikariConfig.addDataSourceProperty("prepStmtCacheSqlLimit", "2048");
			HikariDataSource dataSource = new HikariDataSource(hikariConfig);

			// 创建connection
			// conn = HikariAS400.getDataSource().getConnection();
			conn = dataSource.getConnection();
			stmt = conn.createStatement();
			stmt.execute("set current schema MSRFLIB");

			String createtbl = "DECLARE GLOBAL TEMPORARY TABLE SESSION.TMP2DID("
					+ "    ID     BIGINT GENERATED ALWAYS AS IDENTITY (START WITH 1, INCREMENT BY 1) NOT NULL,"
					+ "    IDHMCD CHAR(15)   default ' ' not null, IDTJYN CHAR(4)    default ' ' not null,"
					+ "    IDEKOT CHAR(5)    default ' ' not null, IDRUID DECIMAL(4) default 0   not null,"
					+ "    IDETNM CHAR(40)   default ' ' not null, CREATETIME TIMESTAMP) ON COMMIT PRESERVE ROWS";
			pstmt = conn.prepareStatement(createtbl, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.execute();

			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.execute();

			sql = "INSERT INTO MZ2DIDP(IDHMCD,IDTJYN,IDEKOT,IDRUID,IDETNM,CREATETIME) "
					+ " SELECT IDHMCD,IDTJYN,IDEKOT,IDRUID,IDETNM,NOW() FROM SESSION.TMP2DID A WHERE NOT EXISTS"
					+ " (SELECT 1 FROM MZ2DIDP B WHERE B.IDHMCD=A.IDHMCD AND B.IDTJYN=A.IDTJYN AND B.IDEKOT=A.IDEKOT)";
			pstmt = conn.prepareStatement(sql, ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY);
			pstmt.execute();

			pstmt.close();
			// 关闭connection
			conn.close();
			dataSource.close();

			log.info("数据已追加，删除文件：" + destFileName);
			// f.delete();
			Path path = Paths.get(destFileName);
			Files.delete(path);

		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return ResultBodyUtil.fail(null);
		} catch (IOException e) {
			e.printStackTrace();
			return ResultBodyUtil.fail(null);
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		return result;
	}

	/**
	 * @param path     指想要下载的文件的路径
	 * @param response
	 * @功能描述 下载文件:将输入流中的数据循环写入到响应输出流中，而不是一次性读取到内存
	 */
	@RequestMapping("/export2did")
	public void export2did(String path, HttpServletResponse response) throws IOException {
		path = "D:\\eclipse-workspace\\mrrbe\\target\\MSRFLIB_MZ2DIDP_TEMPLATE.xlsx";
		// 读到流中
		InputStream inputStream = new FileInputStream(path);// 文件的存放路径
		response.reset();
		response.setContentType("application/octet-stream");
		String filename = new File(path).getName();
		response.addHeader("Content-Disposition", "attachment; filename=" + URLEncoder.encode(filename, "UTF-8"));
		ServletOutputStream outputStream = response.getOutputStream();
		byte[] b = new byte[1024];
		int len;
		// 从输入流中读取一定数量的字节，并将其存储在缓冲区字节数组中，读到末尾返回-1
		while ((len = inputStream.read(b)) > 0) {
			outputStream.write(b, 0, len);
		}
		inputStream.close();
	}

	/**
	 * 文件上传具体实现方法;
	 * 
	 * @param
	 * @return
	 */
	@RequestMapping("/uploadZip")
	@ResponseBody
	public ResultBody fileUpload(
			/* @RequestPart("single") MultipartFile mf, */@RequestPart("multi") MultipartFile[] mfs,
			@RequestPart("data") String dataPath) {
		// System.out.println("单文件上传信息为:" + mf.getOriginalFilename());
		ResultBody result = ResultBodyUtil.success();
		System.out.println("data:" + dataPath);
		File setPath = new File(dataPath);
		String zipPath = "";

		if (dataPath.trim().length() > 0 && dataPath.equalsIgnoreCase("default")) {
			zipPath = "\\\\MCZSVR06\\scan\\设计管理课文件归档";
		} else if (dataPath.trim().length() > 0) {
			if (!setPath.exists()) {
				setPath.mkdir();
				System.out.println("File not exist, create!");
				if (!setPath.exists()) {
					return ResultBodyUtil.fail(ErrorCode.FILE_PATH_NOT_EXIST, "文件路径不存在");
				} else
					zipPath = dataPath;
			} else {
				zipPath = dataPath;
			}
		}

		System.out.println("多文件个数:" + mfs.length + ",output=" + zipPath);
		ApplicationHome h = new ApplicationHome(getClass());
		File jarPath = h.getSource();
		String BasePath = "";
		List<String> listDir = new ArrayList<String>();

		// 1.上传文件
		// 文件夹里文件
		for (MultipartFile m : mfs) {
			// System.out.println("多文件信息:文件名称:" + m.getOriginalFilename() + ",文件大小:" +
			// m.getSize() / 1000 + "kb");
			String fileName = jarPath.getParentFile().toString() + "\\" + m.getOriginalFilename();
			String destFileName = "";
			File f = new File(fileName);
			// 主目录
			if (!f.getParentFile().exists())
				f.getParentFile().mkdirs();
			BasePath = f.getParent();

			String filePreName = FileHelper.getFileName(fileName, Constants.FILE_PRENAME);
			if (FileHelper.getFileName(fileName, Constants.FILE_EXTNAME).equalsIgnoreCase("csv")) {
				destFileName = fileName;
			} else if (filePreName.length() >= 12) {
				// 检测子文件夹是否存在，不存在则创建
				String orgFileName = FileHelper.getFileName(fileName, Constants.FILE_PURENAME);
				filePreName = filePreName.replaceAll("_", "-");
				String subpath = f.getParentFile() + "\\" + filePreName.substring(0, 12);
				File f1 = new File(subpath);
				if (!f1.exists()) {
					f1.mkdirs();
				}
				listDir.add(subpath);
				destFileName = subpath + "\\" + orgFileName;
			}

			try {
				BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream(destFileName));
				// System.out.println(m.getName());
				out.write(m.getBytes());
				out.flush();
				out.close();

			} catch (FileNotFoundException e) {
				e.printStackTrace();
				return ResultBodyUtil.fail(null);
			} catch (IOException e) {
				e.printStackTrace();
				return ResultBodyUtil.fail(null);
			}
		}
		// 2.拷贝公共文件到各文件夹
		File fBase = new File(BasePath);
		File[] fs = fBase.listFiles();
		System.out.println("pre-copy-csvfile..." + BasePath + ", filecnt=" + fs.length);
		for (File f : fs) {
			System.out.println("listdir.0..." + f.getName() + ",f.isfile=" + f.isFile());
			if (f.isFile()) {
				System.out.println("listdir.1..." + f.getName() + ",listDir.size()=" + listDir.size());
				for (int i = 0; i < listDir.size(); i++) {
					System.out.println("listdir.2..." + listDir.get(i));
					FileUtil.copy(BasePath + "\\" + f.getName(), listDir.get(i) + "\\" + f.getName());
				}
				break;
			}
		}
		// 3.压缩
		for (File f : fs) {
			if (f.isDirectory()) {
				System.out.println("========== " + f.getAbsolutePath() + "," + f.getName());
				ZipFileUtil.compressToZip(f.getAbsolutePath(), zipPath, f.getName() + ".zip");
			}
		}

		return result;

//		// 单个文件 
//		if (!mf.isEmpty()) { 
//			try { 
//				BufferedOutputStream out = new BufferedOutputStream(new FileOutputStream("D:\\666.txt"));
//				System.out.println(mf.getName()); 
//				out.write(mf.getBytes()); 
//				out.flush();
//				out.close(); 
//			} catch (FileNotFoundException e) { 
//				e.printStackTrace(); 
//				return "上传失败," + e.getMessage(); 
//			} catch (IOException e) { 
//				e.printStackTrace();
//				return "上传失败," + e.getMessage(); 
//			}
//
//			return "上传成功";
//		} 
//		else { 
//			return "上传失败，因为文件是空的."; 
//		}

	}

	public static void main(String[] args) {
		try {
			String url = "jdbc:postgresql://192.168.137.25:5432/BIDATA";
			Class.forName("org.postgresql.Driver");
			Connection connection = DriverManager.getConnection(url, "postgres", "bi@mcz2023");
			String sql = "select * from device_info";
			Statement statement = connection.createStatement();
			ResultSet rs = statement.executeQuery(sql);
			ResultSetMetaData rsmd = rs.getMetaData();
			int index = 0;
			while (rs.next()) {
				String str = "";
				for (int i = 1; i < rsmd.getColumnCount() + 1; i++) {
					String colName = rsmd.getColumnName(i);
					str += colName + ":" + rs.getString(colName) + ",";
				}
				log.info(index + " " + str);
				index++;
			}

			rs.close();
			statement.close();
			connection.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
