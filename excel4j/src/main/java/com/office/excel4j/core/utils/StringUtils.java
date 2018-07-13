package com.office.excel4j.core.utils;

/**
 * 字符串工具类
 * 
 * @since 2018年6月29日
 * @author 赵凡
 * @version 1.0
 *
 */
public class StringUtils {

	/**
	 * 判断字符串是否为空：
	 * <ul>
	 * <li>""->true</li>
	 * <li>" "->true</li>
	 * <li>null->true</li>
	 * </ul>
	 * 
	 * @param str
	 *            字符串
	 * @return 为空返回true，否则返回false
	 */
	public static boolean isBlank(String str) {
		return str == null || str.trim().equals("");
	}

	/**
	 * 判断字符串是否非空
	 * 
	 * @param str
	 *            字符串
	 * @return 非空返回true，否则返回false
	 * @see StringUtils#isBlank(String)
	 */
	public static boolean isNotBlank(String str) {
		return !isBlank(str);
	}

}
