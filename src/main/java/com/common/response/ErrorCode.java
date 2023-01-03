package com.common.response;

import lombok.Getter;

@Getter
public enum ErrorCode {

    /**
     * 系统全局错误码 SYSTEM_SUCCESS
     * 0、-1、404、500、403
     */
    SUCCESS(0, "success"),
    FAIL(-1, "fail"),
    SYSTEM_ERROR_404(404, "Interface Not Found"),
    SYSTEM_ERROR_500(500, "Interface %s internal error, please contact the administrator"),
    SYSTEM_LANGUAGE_CODE_ERROR(600, "The language encoding is not supported by the system"),
    
    ILLEGAL_PARAMETER(1000, "Illegal parameter"),
    PARAM_NOT_EXIST(1001, "%s does not exist"),
    TOKEN_FAIL(1002, "Token Error"),
    LOGIN_VALIDATE_FAIL(1003, "Validate Fail"),
    FILE_PATH_NOT_EXIST(1004, "File path not exist or No permission!")
    ;
	
    public Integer code;

    public String msg;

    public Integer getCode() {
		return code;
	}

	public void setCode(Integer code) {
		this.code = code;
	}

	public String getMsg() {
		return msg;
	}

	public void setMsg(String msg) {
		this.msg = msg;
	}

	ErrorCode(Integer code, String msg) {
        this.code = code;
        this.msg = msg;
    }
}

