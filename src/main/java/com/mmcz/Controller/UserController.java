package com.mmcz.Controller;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.http.HttpServletRequest;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringBootVersion;
import org.springframework.core.SpringVersion;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.alibaba.fastjson.JSON;
import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;
import com.mmcz.Service.UserService;
import com.mmcz.entity.User;
import com.opslab.helper.CharsetHepler;
import com.opslab.helper.ChineseHelper;

import jcifs.smb.SmbFile;

@RestController
@RequestMapping("/user")
public class UserController {

	@Autowired
	private UserService userService;

	@RequestMapping("/findAll")
//	public List<User> findAll() {
//		return userService.findAll();
//	}
	public ResultBody findAll() {
		List<User> userList = userService.findAll();
		ResultBody result = ResultBodyUtil.success(userList);
		return result;
	}

	@RequestMapping("/getAll")
//	public List<User> getAll() {
//		return userService.getAll();
//	}
	public ResultBody getAll() {
		List<User> userList = userService.getAll();
		ResultBody result = ResultBodyUtil.success(userList);
		return result;
	}

	// json提交
	@RequestMapping("/add")
	public ResultBody add(@RequestBody Map<String, String> data) {
		int type = 1;
		String userid = data.get("userid");
		String username = data.get("username");
		String password = data.get("password");
		String deptid = data.get("deptid");
		String email = data.get("email");

		User user = new User();
		user.setType(type);
		user.setUserid(userid);
		user.setUsername(username);
		user.setPassword(password);
		user.setDeptid(Integer.parseInt(deptid));
		user.setEmail(email);
		int n = userService.save(user);
		System.out.println("dd " + user.getUsername());
		ResultBody result = ResultBodyUtil.success();
		return result;
	}

	// json提交
	@RequestMapping("/addUser")
	public ResultBody addUser(@RequestBody @Validated User user) {// , BindingResult bindingResult
		System.out.println("addUser-in " + JSON.toJSONString(user));
		int type = 1;
		user.setType(type);
		int n = userService.save(user);
		System.out.println("dd " + user.getUsername());

		// 如果有参数校验失败，会将错误信息封装成对象组装在BindingResult里
//		for (ObjectError error : bindingResult.getAllErrors()) {
//			return ResultBodyUtil.fail(error.getDefaultMessage());
//		}
		ResultBody result = ResultBodyUtil.success();
		return result;
	}

	// 表单/json提交
	@RequestMapping("/addxForm")
	public ResultBody addxForm(HttpServletRequest request) {
		Map<String, String> data = new HashMap<String, String>();
		Map<String, String[]> map = request.getParameterMap();
		// 参数名称
		Set<String> key = map.keySet();
		// 参数迭代器
		Iterator<String> iterator = key.iterator();
		while (iterator.hasNext()) {
			String k = iterator.next();
			data.put(k, map.get(k)[0]);
		}

		int type = 1;
		String userid = data.get("userid");
		String username = data.get("username");
		String password = data.get("password");
		String deptid = data.get("deptid");
		String email = data.get("email");

		System.out.println("userid=" + userid + ",username=" + username + ",password=" + password + ",deptid=" + deptid
				+ ",email=" + email);
		User user = new User();
		user.setType(type);
		user.setUserid(userid);
		user.setUsername(username);
		user.setPassword(password);
		user.setDeptid(Integer.parseInt(deptid));
		user.setEmail(email);
		int n = userService.save(user);
		System.out.println("dd " + user.getUsername());
		ResultBody result = ResultBodyUtil.success();
		System.out.println("obj " + JSON.toJSONString(result));
		return result;
	}

	@RequestMapping("/getUserById")
	public ResultBody getUserById(@RequestBody Map<String, String> data) {
		String id = data.get("id");
		User user = userService.getUserById(Integer.parseInt(id));
		System.out.println("dd " + user.getUsername());
		return ResultBodyUtil.success(user);
	}

	@RequestMapping("/getUserByUserId")
	public User getUserByUserId(@RequestBody Map<String, String> data) {
		String userid = data.get("userid");
		User user = userService.getUserByUserId(userid);
		System.out.println("dd " + user.getUsername());
		return user;
	}

	@RequestMapping("/update")
	// JSON parse error: Cannot deserialize value of type `java.lang.String` from
	// Object value (token `JsonToken.START_OBJECT`);
	// @RequestBody 不能用Map<String, String>必须用Map<String, Object>，否则如果参数里有对象就报错，如：
	// {"id":"4","type":"1","userid":"3","username":"倩倩","password":"qq6666","deptid":"2","email":"abc@qq.com","dept":{"deptid":2,"deptname":"调达部"}}
//	public ResultBody update(HttpServletRequest request) {
//		Map<String, String> data = new HashMap<String, String>();
//		Map<String, String[]> map = request.getParameterMap();
//		// 参数名称
//		Set<String> key = map.keySet();
//		// 参数迭代器
//		Iterator<String> iterator = key.iterator();
//		while (iterator.hasNext()) {
//			String k = iterator.next();
//			data.put(k, map.get(k)[0]);
//		}
//
//		Integer id = Integer.parseInt(data.get("id").toString());
//		Integer type = Integer.parseInt(data.get("type").toString());
//		String userid = data.get("userid").toString();
//		String username = data.get("username").toString();
//		String password = data.get("password").toString();
//		String deptid = data.get("deptid").toString();
//		String email = data.get("email").toString();
//
//		User user = new User();
//		user.setId(id);
//		user.setType(type);
//		user.setUserid(userid);
//		user.setUsername(username);
//		user.setPassword(password);
//		user.setDeptid(Integer.parseInt(deptid));
//		user.setEmail(email);
	public ResultBody update(@RequestBody User user) throws UnsupportedEncodingException {
		System.out.println("update-in " + CharsetHepler.toGBK(JSON.toJSONString(user)));
		int n = userService.update(user);
		System.out.println("dd " + CharsetHepler.toGBK(user.getUsername()));
		return ResultBodyUtil.success();
	}

	@RequestMapping("/delete")
	public ResultBody delete(@RequestBody Map<String, String> data) {
		Integer id = Integer.parseInt(data.get("id"));
		userService.delete(id);
		System.out.println("delete " + id);
		return ResultBodyUtil.success();
	}

	@RequestMapping("/fuzzyQuery")
	public ResultBody fuzzyQuery(HttpServletRequest request) {
		Map<String, String> data = new HashMap<String, String>();
		Map<String, String[]> map = request.getParameterMap();
		// 参数名称
		Set<String> key = map.keySet();
		// 参数迭代器
		Iterator<String> iterator = key.iterator();
		while (iterator.hasNext()) {
			String k = iterator.next();
			data.put(k, map.get(k)[0]);
		}

		String text = data.get("text");
		List<User> userList = userService.fuzzyQuery(text);
		return ResultBodyUtil.success(userList);
	}

	@RequestMapping("/login")
	public String login(@RequestBody Map<String, String> data) {
		String username = data.get("username");
		String password = data.get("password");
		if (username.equals("admin@qq.com") && password.equals("123456")) {
			return username;
		}
		return password;
	}

	public static void createFile(String path) {
		InputStream in = null;
		OutputStream out = null;
		try {
//			File srcFile = new File("Z:\\ZP188006\\188003.pdf");
//			File destFile = new File("Z:\\ZP188006\\dest188003.pdf");
			File srcFile = new File(path + "188003.pdf");
			File destFile = new File(path + "dest666-188003.pdf");
			in = new BufferedInputStream(new FileInputStream(srcFile));
			out = new BufferedOutputStream(new FileOutputStream(destFile));
			byte[] buffer = new byte[4096];
			int len = 0; // 读取长度
			while ((len = in.read(buffer, 0, buffer.length)) != -1) {
				out.write(buffer, 0, len);
			}
			out.flush();// 刷新缓冲的输出流
		} catch (Exception e) {
			System.out.println("读取文件发生错误！");
		} finally {
			try {
				if (out != null) {
					out.close();
				}
				if (in != null) {
					in.close();
				}
			} catch (Exception e) {
				System.out.println("关闭资源发生错误！");
			}
		}
	}

	public static void getRemoteFile() {
		InputStream in = null;
		try {
			// 创建远程文件对象
			// smb://用户名:密码/共享的路径/...
			// smb://ip地址/共享的路径/...
			// String remoteUrl = "smb://192.168.xx.xx/file/";
			String remoteUrl = "smb://192.168.121.32/scan/ZP188006/";
			SmbFile remoteFile = new SmbFile(remoteUrl);
			remoteFile.connect();// 尝试连接
			if (remoteFile.exists()) {
				// 获取共享文件夹中文件列表
				SmbFile[] smbFiles = remoteFile.listFiles();
				for (SmbFile smbFile : smbFiles) {
					System.out.println(smbFile.getName());
				}
			} else {
				System.out.println("文件不存在！");
			}
		} catch (Exception e) {
			System.out.println("访问远程文件夹出错：" + e.getLocalizedMessage());
		} finally {
			try {
				if (in != null)
					in.close();
			} catch (Exception e) {
				System.out.println("关闭资源错误");
			}
		}
	}

	public static void main(String[] args) throws Exception {
		Map<String, String> map = new HashMap<String, String>();
		map.put("token", "sdfdfsfd");
		System.out.println(ResultBodyUtil.success(map).toString());
		System.out.println(ResultBodyUtil.paramFail().toString());
		System.out.println(CharsetHepler.toGBK(CharsetHepler.toUTF_8("系统开发课")));
		System.out.println("系统开发课");
		System.out.println(CharsetHepler.toGBK("系统开发课"));
		System.out.println(CharsetHepler.toUTF_8("系统开发课"));
		// 也就是代码页的中文是gbk的，ChineseHelper.getPingYin的参数的中文是需要utf-8的
		System.out.println(ChineseHelper.getPingYin(CharsetHepler.toUTF_8("我顶你的肺")));

		String version = SpringVersion.getVersion();
		String version1 = SpringBootVersion.getVersion();
		System.out.println(version);
		System.out.println(version1);
		createFile("\\\\MCZSVR06\\scan\\ZP188006\\");
	}
}
