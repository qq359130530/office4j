package com.office.excel4j.core;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.office.excel4j.core.exception.AccessExcelException;
import com.office.excel4j.core.utils.StringUtils;

/**
 * Excel文件输出器
 * 
 * @since 2018年6月29日
 * @author 赵凡
 * @version 1.0
 *
 */
public class ExcelWriter {
	
	/**
	 * Excel换号符
	 * 
	 */
	public static final String NEW_LINE = "\n";
	
	// Excel工作簿对象
	private Workbook wb;
	
	// 当前操作的Sheet
	private Sheet curSheet;
	
	/**
	 * 根据Excel版本创建小型Excel输出对象
	 * 
	 * @param version
	 *            Excel版本
	 * @param sheetName
	 *            第一个Sheet名称
	 */
	protected ExcelWriter(Version version, String sheetName) {
		// 根据版本创建工作簿对象
		if (version.equals(Version.XLS)) {
			wb = new HSSFWorkbook();
		} else {
			wb = new XSSFWorkbook();
		}
		
		// 默认初始化一个Sheet
		curSheet = buildSheet(sheetName);
		// 选中当前操作的Sheet
		curSheet.setSelected(true);
	}
	
	/**
	 * 创建大型Excel 2007文件输出器
	 * 
	 * @param sheetName
	 *            第一个Sheet名称
	 * @param rowAccessWindowSize
	 *            每次可操作的行数
	 * @param compress
	 *            是否压缩临时文件
	 */
	protected ExcelWriter(String sheetName, int rowAccessWindowSize, boolean compress) {
		SXSSFWorkbook swb = new SXSSFWorkbook(rowAccessWindowSize);
		swb.setCompressTempFiles(compress);
		wb = swb;
		
		// 默认初始化一个Sheet
		curSheet = buildSheet(sheetName);
		// 选中当前操作的Sheet
		curSheet.setSelected(true);
	}
	
	/**
	 * 根据Excel版本创建小型Excel输出对象
	 * 
	 * @param version
	 *            Excel版本
	 * @return 小型Excel文件输出器
	 */
	public static ExcelWriter getInstance(Version version) {
		return new ExcelWriter(version, null);
	}
	
	/**
	 * 创建大型Excel 2007文件输出器
	 * 
	 * @param sheetName
	 *            第一个Sheet名称
	 * @param rowAccessWindowSize
	 *            每次可操作的行数
	 * @param compress
	 *            是否压缩临时文件
	 * @return 大型Excel 2007文件输出器
	 */
	public static ExcelWriter getInstance(String sheetName, int autoSize, boolean tempFiles) {
		return new ExcelWriter(sheetName, autoSize, tempFiles);
	}
	
	/**
	 * 根据Excel版本创建小型Excel输出对象
	 * 
	 * @param version
	 *            Excel版本
	 * @param sheetName
	 *            第一个Sheet名称
	 * @return 小型Excel文件输出器
	 */
	public static ExcelWriter getInstance(Version version, String sheetName) {
		return new ExcelWriter(version, sheetName);
	}
	
	/**
	 * 获取POI Excel工作簿对象实例
	 * 
	 * @return POI Excel工作簿对象实例
	 */
	public Workbook getWorkbook() {
		return wb;
	}
	
	/**
	 * 将内存中的Excel工作簿对象写入指定的文件中
	 * 
	 * @param desc
	 *            要写入的文件对象
	 */
	public void writeFile(File desc) {
		OutputStream out = null;
		try {
			File parent = desc.getParentFile();
			if (!parent.exists()) {
				parent.mkdirs();
			}
			out = new FileOutputStream(desc);
		} catch (FileNotFoundException e) {
			throw new AccessExcelException("输出Excel文件[ " + desc.getPath() + " ]失败！", e);
		} finally {
			writeStream(out);
		}
	}
	
	/**
	 * 将内存中的工作簿对象写入输出流中
	 * 
	 * @param out
	 *            输出流
	 */
	public void writeStream(OutputStream out) {
		try {
			if (out != null)
				wb.write(out);
		} catch (IOException e) {
			throw new AccessExcelException("输出Excel文件到流中失败！", e);
		} finally {
			if (wb instanceof SXSSFWorkbook) {
				SXSSFWorkbook swb = (SXSSFWorkbook) wb;
				swb.dispose();
			}
			try {
				if (out != null)
					out.close();
			} catch (IOException e) {
				throw new AccessExcelException("关闭Excel输出流失败！", e);
			}
		}
	}
	
	// ====================== Sheet相关操作start =====================
	// ====================== Sheet相关操作     end =====================
	/**
	 * 获取当前Sheet对象
	 * 
	 * @return 当前Sheet对象
	 */
	public Sheet getCurSheet() {
		return curSheet;
	}
	
	/**
	 * 创建Sheet
	 * 
	 * @param sheetName
	 *            Sheet名称
	 * @return Excel文件输出器
	 */
	public ExcelWriter createSheet(String sheetName) {
		buildSheet(sheetName);
		return this;
	}
	
	/**
	 * 创建Sheet并切换为当前操作Sheet
	 * 
	 * @param sheetName
	 *            Sheet名称
	 * @return Excel文件输出器
	 */
	public ExcelWriter createAndSwitchSheet(String sheetName) {
		curSheet = buildSheet(sheetName);
		return this;
	}
	
	/**
	 * 构建一个Sheet对象
	 * 
	 * @param sheetName Sheet名称
	 * @return POI Sheet对象
	 */
	protected Sheet buildSheet(String sheetName) {
		if (StringUtils.isBlank(sheetName)) {
			return wb.createSheet();
		} else {
			return wb.createSheet(WorkbookUtil.createSafeSheetName(sheetName));
		}
	}
	
	/**
	 * 切换当前操作的Sheet
	 * 
	 * @param sheetIndex
	 *            Sheet索引，从1开始
	 * @return Excel文件输出器
	 */
	public ExcelWriter switchSheet(int sheetIndex) {
		curSheet = wb.getSheetAt(sheetIndex - 1);
		return this;
	}
	
	/**
	 * 切换当前操作的Sheet
	 * 
	 * @param sheetName
	 *            Sheet名称
	 * @return Excel文件输出器
	 */
	public ExcelWriter switchSheet(String sheetName) {
		curSheet = wb.getSheet(WorkbookUtil.createSafeSheetName(sheetName));
		return this;
	}
	// ====================== Sheet相关操作     end =====================
	
	// ======================  Row相关操作start  =====================
	
	/**
	 * 获取指定行
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @return POI Row
	 */
	public Row getRow(int rowIdx) {
		int rownum = rowIdx - 1;
		Row row = curSheet.getRow(rownum);
		if (row == null) {
			row = curSheet.createRow(rownum);
		}
		return row;
	}
	
	// ======================  Row相关操作     end  =====================
	
	// ======================  Cell相关操作start =====================
	
	/**
	 * 获取单元格
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @return POI Cell
	 */
	public Cell getCell(int rowIdx, String colStr) {
		int colIdx = CellReference.convertColStringToIndex(colStr);
		return getCell(rowIdx, colIdx);
	}

	/**
	 * 获取单元格
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colIdx
	 *            列号，从0开始
	 * @return POI Cell
	 */
	private Cell getCell(int rowIdx, int colIdx) {
		Row row = getRow(rowIdx);
		Cell cell = row.getCell(colIdx);
		if (cell == null) {
			cell = row.createCell(colIdx);
		}
		return cell;
	}
	
	/**
	 * 合并单元格
	 * 
	 * @param ref
	 *            要合并的单元格范围引用，例如：A3:B5
	 * @return Excel文件输出器
	 */
	public ExcelWriter mergeCell(String ref) {
		curSheet.addMergedRegion(CellRangeAddress.valueOf(ref));
		return this;
	}
	
	// ======================  Cell相关操作     end =====================
	
	// ======================  输出相关操作start  =====================
	
	/**
	 * 输出字符串值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeString(int rowIdx, String colStr, String value) {
		if (StringUtils.isNotBlank(value)) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.STRING);
			cell.setCellValue(value);
		}
		return this;
	}
	
	/**
	 * 输出Double值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeNumber(int rowIdx, String colStr, Number value) {
		if (value != null) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.NUMERIC);
			if (value instanceof BigDecimal) {// BigDecimal
				cell.setCellValue(value.toString());
			} else {// 其它数值类型
				cell.setCellValue(value.doubleValue());
			}
		}
		return this;
	}
	
	/**
	 * 输出日期值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeDate(int rowIdx, String colStr, Date value) {
		if (value != null) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.NUMERIC);
			cell.setCellValue(value);
		}
		return this;
	}
	
	/**
	 * 输出日期值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeCalendar(int rowIdx, String colStr, Calendar value) {
		if (value != null) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.NUMERIC);
			cell.setCellValue(value);
		}
		return this;
	}
	
	/**
	 * 输出Boolean值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeBoolean(int rowIdx, String colStr, Boolean value) {
		if (value != null) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.BOOLEAN);
			cell.setCellValue(value);
		}
		return this;
	}
	
	/**
	 * 输出公式
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeFormula(int rowIdx, String colStr, String value) {
		if (value != null) {
			Cell cell = getCell(rowIdx, colStr);
			cell.setCellType(CellType.FORMULA);
			cell.setCellValue(value);
		}
		return this;
	}
	
	/**
	 * 输出Object值
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param value
	 *            要输出的值
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeObject(int rowIdx, String colStr, Object value) {
		if (value != null) {
			if (value instanceof String) {
				writeString(rowIdx, colStr, (String) value);
			} else if (value instanceof Number) {
				writeNumber(rowIdx, colStr, (Number) value);
			} else if (value instanceof Date) {
				writeDate(rowIdx, colStr, (Date) value);
			} else if (value instanceof Calendar) {
				writeCalendar(rowIdx, colStr, (Calendar) value);
			} else if (value instanceof Boolean) {
				writeBoolean(rowIdx, colStr, (Boolean) value);
			} else {// 不支持的格式，调用对象的toString()
				writeString(rowIdx, colStr, value.toString());
			}
		}
		return this;
	}
	
	/**
	 * 输出一行输出
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param firstColStr
	 *            第一列列号，从A开始
	 * @param values
	 *            值列表
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeRow(int rowIdx, String firstColStr, Object... values) {
		if (values != null && values.length > 0) {
			int curColIdx = CellReference.convertColStringToIndex(firstColStr);
			for (Object value : values) {
				writeObject(rowIdx, CellReference.convertNumToColString(curColIdx), value);
				curColIdx ++;
			}
		}
		return this;
	}
	
	/**
	 * 输出一块内容
	 * 
	 * @param firstRowIdx
	 *            开始行号，从1开始
	 * @param firstColStr
	 *            开始列号，从A开始
	 * @param rows
	 *            要写入的集合
	 * @param rowMapper
	 *            行映射器
	 * @return Excel文件输出器
	 */
	public <T> ExcelWriter writeBlock(int firstRowIdx, String firstColStr, List<T> rows, RowMapper<T> rowMapper) {
		if (rows != null && !rows.isEmpty()) {
			for (int i = 0; i < rows.size(); i++) {
				Object result = rowMapper.map(i + 1, rows.get(i));
				if (result != null) {
					Object[] objs = null;
					if (result.getClass().isArray()) {
						objs = (Object[]) result;
					} else {
						objs = new Object[] { result };
					}
					writeRow(firstRowIdx + i, firstColStr, objs);
				}
			}
		}
		return this;
	}
	
	/**
	 * 从A列开始，输出一块内容
	 * 
	 * @param firstRowIdx
	 *            开始行号，从1开始
	 * @param rows
	 *            要写入的集合
	 * @param rowMapper
	 *            行映射器
	 * @return Excel文件输出器
	 */
	public <T> ExcelWriter writeBlock(int firstRowIdx, List<T> rows, RowMapper<T> rowMapper) {
		writeBlock(firstRowIdx, "A", rows, rowMapper);
		return this;
	} 
	
	/**
	 * 输出图片
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param imageFile
	 *            图片文件
	 * @return Excel文件输出器
	 */
	public ExcelWriter writeImage(int rowIdx, String colStr, File imageFile) {
		if (imageFile != null && imageFile.exists() && imageFile.isFile()) {
			String pictureName = imageFile.getName().toUpperCase();
			try (InputStream input = new FileInputStream(imageFile)) {
				// 计算列索引
				int colIdx = CellReference.convertColStringToIndex(colStr);
				
				// 添加图片到工作簿
				byte[] bytes = IOUtils.toByteArray(input);
				Integer pictureIdx = null;
				if (pictureName.endsWith(".PNG")) {
					pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
				} else if (pictureName.endsWith(".JPEG") || pictureName.endsWith(".JPG")) {
					pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
				} else {
					throw new AccessExcelException("不支持图片格式[" + pictureName + "]！");
				}
				
				// 创建画布
				CreationHelper helper = wb.getCreationHelper();
				Drawing drawing = curSheet.createDrawingPatriarch();
				// 设置图片单元格位置
				ClientAnchor anchor = helper.createClientAnchor();
				anchor.setCol1(colIdx);
			    anchor.setRow1(rowIdx - 1);
			    
			    // 绘制图片
			    Picture pict = drawing.createPicture(anchor, pictureIdx);
			    // 按比例自动调整图片大小
			    pict.resize(1.0);
			} catch (IOException e) {
				// 抛出根异常
				throw new AccessExcelException("添加图片[" + pictureName + "]到Excel失败！", e);
			}
		}
		return this;
	}
	
	// ======================  输出相关操作     end  =====================
	
	// ======================  样式相关操作start  =====================
	
	/**
	 * 绑定单元格样式
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param colStr
	 *            列号，从A开始
	 * @param cellStyle
	 *            单元格样式
	 * @return Excel文件输出器
	 */
	public ExcelWriter bindCellStyle(int rowIdx, String colStr, CellStyle cellStyle) {
		Cell cell = getCell(rowIdx, colStr);
		cell.setCellStyle(cellStyle);
		return this;
	}
	
	/**
	 * 绑定行样式
	 * 
	 * @param rowIdx
	 *            行号，从1开始
	 * @param startColStr
	 *            第一列列号，从A开始
	 * @param colCount
	 *            总列数
	 * @param cellStyle
	 *            单元格样式
	 * @return Excel文件输出器
	 */
	public ExcelWriter bindRowStyle(int rowIdx, String startColStr, int colCount, CellStyle cellStyle) {
		for (int x = 0; x < colCount; x++) {
			int colIdx = CellReference.convertColStringToIndex(startColStr);
			Cell cell = getCell(rowIdx, colIdx + x);
			cell.setCellStyle(cellStyle);
		}
		return this;
	}
	
	/**
	 * 绑定块样式
	 * 
	 * @param startRowIdx
	 *            第一行行号，从1开始
	 * @param rowCount
	 *            总行数
	 * @param startColStr
	 *            第一列列号，从A开始
	 * @param colCount
	 *            总列数
	 * @param cellStyle
	 *            单元格样式
	 * @return Excel文件输出器
	 */
	public ExcelWriter bindBlockStyle(int startRowIdx, int rowCount, String startColStr, int colCount, CellStyle cellStyle) {
		for (int y = 0; y < rowCount; y++) {
			for (int x = 0; x < colCount; x++) {
				int colIdx = CellReference.convertColStringToIndex(startColStr);
				Cell cell = getCell(startRowIdx + y, colIdx + x);
				cell.setCellStyle(cellStyle);
			}
		}
		return this;
	}
	
	/**
	 * 创建对齐样式
	 * 
	 * @param halign
	 *            垂直对齐
	 * @param valign
	 *            水平对齐
	 * @return 对齐样式
	 */
	public CellStyle createAlignStyle(HorizontalAlignment halign, VerticalAlignment valign) {
		CellStyle style = wb.createCellStyle();
		if (halign != null)
			style.setAlignment(halign);
		if (valign != null)
			style.setVerticalAlignment(valign);
		return style;
	}
	
	/**
	 * 创建单元格边框样式
	 * 
	 * @return 单元格边框样式
	 */
	public CellStyle createBorderStyle() {
	    CellStyle style = wb.createCellStyle();
	    style.setBorderBottom(BorderStyle.THIN);
	    style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderLeft(BorderStyle.THIN);
	    style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderRight(BorderStyle.THIN);
	    style.setRightBorderColor(IndexedColors.BLACK.getIndex());
	    style.setBorderTop(BorderStyle.THIN);
	    style.setTopBorderColor(IndexedColors.BLACK.getIndex());
	    return style;
	}
	
	/**
	 * 创建背景色样式
	 * 
	 * @param bg
	 *            背景颜色
	 * @see IndexedColors
	 * @return 背景色样式
	 */
	public CellStyle createBackgroundColorStyle(short bg) {
		CellStyle style = wb.createCellStyle();
		style.setFillBackgroundColor(bg);// IndexedColors
		return style;
	}
	
	/**
	 * 创建字体
	 * 
	 * @param fontHeightInPoints
	 *            字体高度
	 * @param fontName
	 *            字体名称
	 * @param bold
	 *            是否加粗
	 * @param italic
	 *            是否倾斜
	 * @return POI Font
	 */
	public Font createFont(short fontHeightInPoints, String fontName, boolean bold, boolean italic) {
		Font font = wb.createFont();
	    font.setFontHeightInPoints(fontHeightInPoints);
	    font.setFontName(fontName);
	    font.setItalic(italic);
	    font.setBold(bold);
		return font;
	}
	
	/**
	 * 创建字体样式
	 * 
	 * @param fontHeightInPoints
	 *            字体高度
	 * @param fontName
	 *            字体名称
	 * @param bold
	 *            是否加粗
	 * @param italic
	 *            是否倾斜
	 * @return 字体样式
	 */
	public CellStyle createFontStyle(short fontHeightInPoints, String fontName, boolean bold, boolean italic) {
		Font font = createFont(fontHeightInPoints, fontName, bold, italic);
	    CellStyle style = wb.createCellStyle();
	    style.setFont(font);
	    return style;
	}
	
	/**
	 * 创建数据格式化器
	 * 
	 * @param format
	 *            数据格式
	 * @return 数据格式化器
	 */
	public short createDataFormat(String format) {
		DataFormat dataFormat = wb.createDataFormat();
		return dataFormat.getFormat(format);
	}
	
	/**
	 * 创建数据格式化样式
	 * 
	 * @param format
	 *            数据格式
	 * @return 数据格式化样式
	 */
	public CellStyle createDataFormatStyle(String format) {
		short fmt = createDataFormat(format);
		CellStyle style = wb.createCellStyle();
		style.setDataFormat(fmt);
		return style;
	}
	
	/**
	 * 设置冻结窗口
	 * 
	 * @param rowSplit
	 *            冻结行号，第四象限最近原点的单元格，从1开始
	 * @param colSplit
	 *            冻结列号，第四象限最近原点的单元格，从A开始
	 * @return Excel文件输出器
	 */
	public ExcelWriter freezePane(int rowSplit, String colSplit) {
		curSheet.createFreezePane(CellReference.convertColStringToIndex(colSplit), rowSplit - 1);
		return this;
	}
	
	/**
	 * 设置默认行高
	 * 
	 * @param width
	 *            行高，单位为榜
	 * @return Excel文件输出器
	 */
	public ExcelWriter setDefaultRowHeight(short height) {
		curSheet.setDefaultRowHeight((short)(height * 20));
		return this;
	}
	
	/**
	 * 设置行高
	 * 
	 * @param height
	 *            行高，单位为榜
	 * @param rowIdxs
	 *            行号列表
	 * @return Excel文件输出器
	 */
	public ExcelWriter setRowHeight(int height, int... rowIdxs) {
		if (rowIdxs != null && rowIdxs.length > 0) {
			for (int rowIdx : rowIdxs) {
				Row row = getRow(rowIdx);
				row.setHeightInPoints(height);
			}
		}
		return this;
	}
	
	/**
	 * 设置默认列宽
	 * 
	 * @param width
	 *            列度，单位一个字符宽度
	 * @return Excel文件输出器
	 */
	public ExcelWriter setDefaultColumnWidth(int width) {
		curSheet.setDefaultColumnWidth(width);
		return this;
	}
	
	/**
	 * 设置列宽
	 * 
	 * @param width
	 *            列度，单位一个字符宽度
	 * @param colStrs
	 *            列号列表
	 * @return Excel文件输出器
	 */
	public ExcelWriter setColumnWidth(int width, String... colStrs) {
		if (colStrs != null && colStrs.length > 0) {
			for (String colStr : colStrs) {
				curSheet.setColumnWidth(CellReference.convertColStringToIndex(colStr), width * 256);
			}
		}
		return this;
	}
	
	// ======================  样式相关操作     end  =====================
	
}
