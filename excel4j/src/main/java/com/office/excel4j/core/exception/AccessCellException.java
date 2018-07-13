package com.office.excel4j.core.exception;

/**
 * 访问单元格异常
 * 
 * @since 2018年7月3日
 * @author 赵凡
 * @version 1.0
 *
 */
public class AccessCellException extends RuntimeException {

	/**
	 * 
	 */
	private static final long serialVersionUID = -6275461817116886287L;

	private int sheetIdx;// sheet索引，从1开始

	private String sheetName;// sheet名称

	private int rowIdx;// 行索引，从1开始

	private String colStr;// 列索引，从A开始

	/**
	 * 创建访问单元格异常
	 * 
	 * @param sheetIdx
	 *            sheet索引，从1开始
	 * @param sheetName
	 *            sheet名称
	 * @param rowIdx
	 *            行索引，从1开始
	 * @param colStr
	 *            列索引，从A开始
	 */
	public AccessCellException(int sheetIdx, String sheetName, int rowIdx, String colStr) {
		this(sheetIdx, sheetName, rowIdx, colStr, "处理异常！");
	}

	/**
	 * 创建访问单元格异常
	 * 
	 * @param sheetIdx
	 *            sheet索引，从1开始
	 * @param sheetName
	 *            sheet名称
	 * @param rowIdx
	 *            行索引，从1开始
	 * @param colStr
	 *            列索引，从A开始
	 * @param message
	 *            异常消息
	 */
	public AccessCellException(int sheetIdx, String sheetName, int rowIdx, String colStr, String message) {
		super("Sheet[index=" + sheetIdx + ",name=" + sheetName + "]下的单元格[row=" + rowIdx + ",col=" + colStr + "]:"
				+ message);
		this.sheetIdx = sheetIdx;
		this.sheetName = sheetName;
		this.rowIdx = rowIdx;
		this.colStr = colStr;
	}

	/**
	 * 获取sheet索引，从1开始
	 * 
	 * @return sheet索引，从1开始
	 */
	public int getSheetIdx() {
		return sheetIdx;
	}

	/**
	 * 获取sheet名称
	 * 
	 * @return sheet名称
	 */
	public String getSheetName() {
		return sheetName;
	}

	/**
	 * 获取行索引，从1开始
	 * 
	 * @return 行索引，从1开始
	 */
	public int getRowIdx() {
		return rowIdx;
	}

	/**
	 * 获取列索引，从A开始
	 * 
	 * @return 列索引，从A开始
	 */
	public String getColStr() {
		return colStr;
	}

}
