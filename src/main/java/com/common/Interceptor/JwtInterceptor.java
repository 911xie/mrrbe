package com.common.Interceptor;

import java.io.IOException;
import java.io.PrintWriter;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.method.HandlerMethod;
import org.springframework.web.servlet.HandlerInterceptor;

import com.auth0.jwt.exceptions.JWTDecodeException;
import com.auth0.jwt.exceptions.JWTVerificationException;
import com.common.exception.ServiceException;
import com.common.response.ErrorCode;
import com.common.response.ResultBodyUtil;
import com.mmcz.Service.UserService;
import com.mmcz.entity.User;
import com.opslab.helper.JWTUtil;

public class JwtInterceptor implements HandlerInterceptor {
	@Autowired
	private UserService userService;

	public void ReturnRsp(HttpServletResponse response, String ErrMsg) throws IOException {
		// 重置response
		response.reset();
		// 设置编码格式
		response.setCharacterEncoding("UTF-8");
		response.setContentType("application/json;charset=UTF-8");

		PrintWriter pw = response.getWriter();
		pw.write(ResultBodyUtil.fail(ErrorCode.TOKEN_FAIL, ErrMsg).toString());

		pw.flush();
		pw.close();
	}

	@Override
	public boolean preHandle(HttpServletRequest request, HttpServletResponse response, Object handler)
			throws IOException {
		System.out.println("preHandle...0");
		String token = request.getHeader("token");
		// 如果不是映射到方法直接通过
		if (!(handler instanceof HandlerMethod)) {
			return true;
		}

		// 执行认证
		System.out.println("preHandle...11");
		if (StringUtils.isBlank(token) || (token == null)) {
			throw new ServiceException(ErrorCode.TOKEN_FAIL.getCode(), "无token，请重新登录");
		}
		System.out.println("preHandle...22 " + token);
		// 获取 token 中的 userid
		String userId;
		try {
			userId = JWTUtil.getAudience(token);
			System.out.println("getAudience..." + userId);
		} catch (JWTDecodeException j) {
			throw new ServiceException(ErrorCode.TOKEN_FAIL.getCode(), "token验证失败,请重新登录");
		}

		// 根据token中的userid查询数据库*3. N/*
		User user = userService.getUserByUserId(userId);
		if (user == null) {
			throw new ServiceException(ErrorCode.TOKEN_FAIL.getCode(), "用户不存在,请重新登录");
		}

		if (JWTUtil.isExpire(token)) {
			throw new ServiceException(ErrorCode.TOKEN_FAIL.getCode(), "Token已超时，请重新登录");
		}

		try {
			// 验证 token
			JWTUtil.verifyToken(token, user.getUserid());
		} catch (JWTVerificationException e) {
			// ReturnRsp(response, "token验证失败,请重新登录666");
			throw new ServiceException(ErrorCode.TOKEN_FAIL.getCode(), "token验证失败,请重新登录");
		}
		return true;
	}
}
