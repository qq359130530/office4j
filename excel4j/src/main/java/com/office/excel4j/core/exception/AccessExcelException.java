package com.office.excel4j.core.exception;

import java.util.ArrayList;
import java.util.List;

import com.office.excel4j.core.Cell;

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

	// 访问单元格异常列表
	private List<AccessCellException> accessCessExceptions = new ArrayList<>();

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

	/**
	 * 添加访问单元格异常
	 * 
	 * @param e
	 *            访问单元格异常
	 * @return 访问Excel异常
	 */
	public AccessExcelException addCellException(AccessCellException e) {
		accessCessExceptions.add(e);
		return this;
	}

	/**
	 * 添加访问单元格异常
	 * 
	 * @param cell
	 *            异常单元格
	 * @param message
	 *            异常消息
	 * @return 访问Excel异常
	 */
	public AccessExcelException addCellException(Cell cell, String message) {
		accessCessExceptions.add(new AccessCellException(cell.getSheetIdx(), cell.getSheetName(), cell.getRowIdx(),
				cell.getColStr(), message));
		return this;
	}

	/**
	 * 获取访问单元格异常列表
	 * 
	 * @return 访问单元格异常列表
	 */
	public List<AccessCellException> getAccessCessExceptions() {
		return accessCessExceptions;
	}

}
