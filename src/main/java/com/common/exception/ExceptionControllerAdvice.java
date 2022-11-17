package com.common.exception;

import org.springframework.validation.ObjectError;
import org.springframework.web.bind.MethodArgumentNotValidException;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.RestControllerAdvice;

import com.common.response.ErrorCode;
import com.common.response.ResultBody;
import com.common.response.ResultBodyUtil;

@RestControllerAdvice
//@ControllerAdvice
public class ExceptionControllerAdvice {

	@ExceptionHandler(MethodArgumentNotValidException.class)
	public ResultBody MethodArgumentNotValidExceptionHandler(MethodArgumentNotValidException e) {
		System.out.println("ExceptionControllerAdvice...........");
		// 从异常对象中拿到ObjectError对象
		ObjectError objectError = e.getBindingResult().getAllErrors().get(0);
		// 然后提取错误提示信息进行返回
		return ResultBodyUtil.fail(objectError.getDefaultMessage());// objectError.getDefaultMessage();
	}

	// 统一处理全局异常（Exception），返回统一数据给前端 也可以自定义需要处理的异常
	// 参数是你需要捕获的异常类
	@ExceptionHandler({ ServiceException.class })
	public ResultBody exceptionHandler(Exception e) {
		// Result是统一封装的一个返回实体
		System.out.println("ServiceException~~~~~");
		return ResultBodyUtil.fail(ErrorCode.TOKEN_FAIL, e.getMessage());
	}
}
