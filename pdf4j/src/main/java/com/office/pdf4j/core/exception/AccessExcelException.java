package com.office.pdf4j.core.exception;

/**
 * 访问Excel异常
 * 
 * @since 2018年7月3日
 * @author 赵凡
 * @version 1.0
 *
 */
public class AccessExcelException extends RuntimeException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 105908678931565982L;

	/**
	 * 创建访问Excel异常
	 * 
	 * @param message
	 *            异常消息
	 * @param cause
	 *            异常堆栈信息
	 */
	public AccessExcelException(String message, Throwable cause) {
		super(message, cause);
	}

	/**
	 * 创建访问Excel异常
	 * 
	 * @param message
	 *            异常消息
	 */
	public AccessExcelException(String message) {
		this(message, null);
	}

	/**
	 * 创建访问Excel异常
	 * 
	 * @param cause
	 *            异常堆栈信息
	 */
	public AccessExcelException(Throwable cause) {
		this("访问Excel失败！", null);
	}

}
