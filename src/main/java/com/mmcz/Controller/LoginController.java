package com.mmcz.Controller;

import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import com.alibaba.fastjson.JSONArray;
import com.alibaba.fastjson.JSONObject;
import com.common.response.ErrorCode;
import com.mmcz.Service.UserService;
import com.mmcz.entity.User;
import com.opslab.helper.JWTUtil;

@RestController
@RequestMapping("/user")
public class LoginController {
	@Autowired
	private UserService userService;

	@Autowired
	private JavaMailSender mailSender;

	@Value("${spring.mail.username}")
	private String emailFrom;

	/*
	 * // 登录功能实现
	 * 
	 * @RequestMapping(value = "/login", method = RequestMethod.POST) public
	 * JSONObject login(@RequestBody Map<String, String> data) { String userid =
	 * data.get("userid"); String username = data.get("username"); String usernick =
	 * data.get("usernick"); String email = data.get("email"); String password =
	 * data.get("password"); System.out.println("/user/login113 " + userid + "," +
	 * password); System.out.println("-------------001"); User user =
	 * userService.findByIdAndPwd(userid, password);
	 * System.out.println("-------------002"); JSONObject result = new JSONObject();
	 * if (user == null) { result.put("state", 400); } else { // String token =
	 * TokenUtil.sign("userName"); String token = JWTUtil.generateToken(userid,
	 * username, email); System.out.println(token); result.put("state", 200);
	 * result.put("token", token); } System.out.println("-------------0031"); return
	 * result; }
	 */
	// 登录功能实现
	@RequestMapping(value = "/login", method = RequestMethod.POST)
//	public JSONObject login(@RequestParam(value = "userid", required = false) String userid,
//			@RequestParam(value = "password", required = false) String password) {
	public JSONObject login(@RequestBody Map<String, String> data) {
		String userid = data.get("userid");
		String username = data.get("username");
		String usernick = data.get("usernick");
		String email = data.get("email");
		String password = data.get("password");
		/*
		 * // 创建简单邮件消息对象 SimpleMailMessage message = new SimpleMailMessage();
		 * System.out.println("..." + emailFrom); // 设置发件人 message.setFrom(emailFrom);
		 * // 设置收件人 message.setTo("303843275@qq.com"); // 设置邮件主题
		 * message.setSubject("邮件测试"); // 设置邮件内容
		 * message.setText("这个是Springboot发出的简单邮件"); // 发送邮件 mailSender.send(message);
		 */
		System.out.println("/user/login113 " + userid + "," + password);
		System.out.println("-------------001");
		User user = userService.findByIdAndPwd(userid, password);
		System.out.println("-------------002");
		String token = "";
		JSONObject result = new JSONObject();
		if (user == null) {
			result.put("code", ErrorCode.LOGIN_VALIDATE_FAIL.getCode());
			result.put("msg", ErrorCode.LOGIN_VALIDATE_FAIL.getMsg());
		} else {
			// String token = TokenUtil.sign("userName");
			token = JWTUtil.generateToken(userid, username, email);
			System.out.println(token);
			result.put("code", ErrorCode.SUCCESS.getCode());
			result.put("msg", ErrorCode.SUCCESS.getMsg());
			JSONObject retData = new JSONObject();
			JSONArray router = new JSONArray();
			router.add("user");
			router.add("dept");

			retData.put("token", token);
			retData.put("routerList", router);
			result.put("data", retData);
		}
		System.out.println("-------------0031");
		return result;
	}

}
