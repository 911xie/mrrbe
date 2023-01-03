package com.mmcz.Controller;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.springframework.boot.system.ApplicationHome;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestPart;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.alibaba.fastjson.JSONObject;
import com.common.Constants;
import com.common.response.ErrorCode;
import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;
import com.opslab.helper.FileHelper;
import com.opslab.helper.HikariAS400;
import com.opslab.helper.HikariMsSql;
import com.opslab.helper.HikariMySql;
import com.opslab.helper.ZipFileUtil;
import com.opslab.util.FileUtil;

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

	/**
	 * 文件上传具体实现方法;
	 * 
	 * @param file
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
				return ResultBodyUtil.fail(ErrorCode.FILE_PATH_NOT_EXIST, "文件路径不存在");
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
			// System.out.println("Basepath..." + BasePath);
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
					listDir.add(subpath);
				}
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
		for (File f : fs) {
			if (f.isFile()) {
				for (int i = 0; i < listDir.size(); i++) {
					System.out.println("listdir..." + listDir.get(i));
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

}
