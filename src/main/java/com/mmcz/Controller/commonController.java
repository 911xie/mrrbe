package com.mmcz.Controller;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.Map;

import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.alibaba.fastjson.JSONObject;
import com.opslab.helper.HikariAS400;
import com.opslab.helper.HikariMsSql;
import com.opslab.helper.HikariMySql;

import net.sf.json.JSONArray;

@RestController
@RequestMapping("/common")
public class commonController {
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

	// 对账查询Reconcile
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

}
