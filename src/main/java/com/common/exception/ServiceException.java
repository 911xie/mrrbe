package com.common.exception;

public class ServiceException extends RuntimeException {
	private int code = 500;

	public ServiceException() {
	}

	public ServiceException(String message) {
		super(message);
	}

	public ServiceException(String message, Throwable cause) {
		super(message, cause);
	}

	public ServiceException(Throwable cause) {
		super(cause);
	}

	public ServiceException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	public ServiceException(int code) {
		this.code = code;
	}

	public ServiceException(int code, String message) {
		super(message);
		this.code = code;
	}

	public ServiceException(int code, String message, Throwable cause) {
		super(message, cause);
		this.code = code;
	}

	public ServiceException(int code, Throwable cause) {
		super(cause);
		this.code = code;
	}

	public ServiceException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace,
			int code) {
		super(message, cause, enableSuppression, writableStackTrace);
		this.code = code;
	}

	public int getCode() {
		return code;
	}
}
