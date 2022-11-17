package com.common.response;

public class ResultBodyUtil {

	public static ResultBody success() {
		return resultBody(ErrorCode.SUCCESS);
	}

	public static ResultBody success(Object data) {
		return resultBody(ErrorCode.SUCCESS, data);
	}

	public static ResultBody fail(Object data) {
		return resultBody(ErrorCode.FAIL, data);
	}

	public static ResultBody fail(ErrorCode errorCode, Object data) {
		return resultBody(errorCode, data);
	}

	public static ResultBody resultBody(ErrorCode errorCode, String... filedsName) {
		return new ResultBody(errorCode, filedsName);
	}

	public static ResultBody resultBody(ErrorCode errorCode, Object data, String... filedsName) {
		return new ResultBody(errorCode, data, filedsName);
	}

	public static ResultBody paramFail() {
		return resultBody(ErrorCode.ILLEGAL_PARAMETER);
	}
}