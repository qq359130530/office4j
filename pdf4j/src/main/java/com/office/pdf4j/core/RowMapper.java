package com.office.pdf4j.core;

/**
 * 行映射器
 * 
 * @since 2018年7月2日
 * @author 赵凡
 * @version 1.0
 *
 */
public interface RowMapper<T> {

	/**
	 * 行映射
	 * 
	 * @param rowObject
	 *            rows中的实体对象
	 * @return 处理结果
	 */
	public String[] map(T rowObject);

}
