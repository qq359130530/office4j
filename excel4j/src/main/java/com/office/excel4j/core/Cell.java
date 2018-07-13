package com.office.excel4j.core;

import java.util.Calendar;
import java.util.Date;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;

import com.office.excel4j.core.exception.AccessExcelException;

/**
 * 单元格
 * 
 * @since 2018年7月3日
 * @author 赵凡
 * @version 1.0
 *
 */
public class Cell {

	private int sheetIdx;// sheet索引，从1开始
	
	private String sheetName;// sheet名称
	
	private int rowIdx;// 行索引，从1开始
	
	private String colStr;// 列索引，从A开始
	
	private Object value;// 值

	/**
	 * 创建单元格
	 * 
	 * @param sheetIdx
	 *            sheet索引，从1开始
	 * @param sheetName
	 *            sheet名称
	 * @param rowIdx
	 *            行索引，从1开始
	 * @param colIdx
	 *            列索引，从A开始
	 * @param value
	 *            值
	 */
	public Cell(int sheetIdx, String sheetName, int rowIdx, int colIdx, Object value) {
		this.sheetIdx = sheetIdx;
		this.sheetName = sheetName;
		this.rowIdx = rowIdx + 1;
		this.colStr = CellReference.convertNumToColString(colIdx);
		this.value = value;
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

	/**
	 * 获取字符串类型值
	 * 
	 * @return 单元格值
	 * @throws AccessExcelException
	 *             单元格类型是非字符串类型
	 */
	public String getStringValue() throws AccessExcelException {
		if (value != null) {
			if (value instanceof String) {
				return value.toString();
			} else {
				throw new AccessExcelException("值[" + value + "]不是字符串类型！");
			}
		}
		return null;
	}
	
	/**
	 * 获取数值类型值
	 * 
	 * @return 单元格值
	 * @throws AccessExcelException
	 *             单元格类型是非数值类型
	 */
	public Number getNumberValue() throws AccessExcelException {
		if (value != null) {
			if (value instanceof Number) {
				return (Number) value;
			} else {
				throw new AccessExcelException("值[" + value + "]不是数值类型！");
			}
		}
		return null;
	}
	
	/**
	 * 获取日期类型值
	 * 
	 * @return 单元格值
	 * @throws AccessExcelException
	 *             单元格类型是非日期类型
	 */
	public Date getDateValue() throws AccessExcelException {
		if (value != null) {
			if (value instanceof Number) {
				return DateUtil.getJavaDate(((Number) value).doubleValue());
			} else {
				throw new AccessExcelException("值[" + value + "]不是日期类型！");
			}
		}
		return null;
	}
	
	/**
	 * 获取日期类型值
	 * 
	 * @return 单元格值
	 * @throws AccessExcelException
	 *             单元格类型是非日期类型
	 */
	public Calendar getCalendarValue() throws AccessExcelException {
		if (value != null) {
			if (value instanceof Number) {
				return DateUtil.getJavaCalendar(((Number) value).doubleValue());
			} else {
				throw new AccessExcelException("值[" + value + "]不是日期类型！");
			}
		}
		return null;
	}
	
	/**
	 * 获取布尔类型值
	 * 
	 * @return 单元格值
	 * @throws AccessExcelException
	 *             单元格类型是非布尔类型
	 */
	public Boolean getBooleanValue() throws AccessExcelException {
		if (value != null) {
			if (value instanceof String) {
				return Boolean.valueOf(value.toString());
			} else {
				throw new AccessExcelException("值[" + value + "]不是布尔类型！");
			}
		}
		return null;
	}
	
}
