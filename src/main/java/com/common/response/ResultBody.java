package com.common.response;

import java.io.Serializable;

import com.fasterxml.jackson.annotation.JsonInclude;

import lombok.Getter;

//<dependency>
//<groupId>org.projectlombok</groupId>
//<artifactId>lombok</artifactId>
//<version>1.18.4</version>
//<scope>provided</scope>
//</dependency>

@Getter
public class ResultBody implements Serializable {

	private Integer code;

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

	public Object getData() {
		return data;
	}

	private String msg;

	@JsonInclude(JsonInclude.Include.NON_NULL)
	private Object data;

	public ResultBody(ErrorCode code, String... filedsName) {
		this.code = code.getCode();
		this.msg = String.format(code.getMsg(), filedsName);
	}

	public ResultBody(ErrorCode code, Object data, String... filedsName) {
		this.code = code.getCode();
		this.msg = String.format(code.getMsg(), filedsName);
		this.data = data;
	}

	public ResultBody(Integer code, String msg, Object data) {
		this.code = code;
		this.msg = msg;
		this.data = data;
	}

	public ResultBody(Integer code, String msg) {
		this.code = code;
		this.msg = msg;
	}

	public ResultBody() {
	}

	public ResultBody setData(Object data) {
		this.data = data;
		return this;
	}

	public String toString() {
		return "code=" + this.code + "\r\n" + this.msg + "\r\n" + this.data;

	}

}