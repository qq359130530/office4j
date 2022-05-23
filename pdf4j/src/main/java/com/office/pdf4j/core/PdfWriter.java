package com.office.pdf4j.core;

import com.office.pdf4j.core.exception.AccessExcelException;
import com.spire.pdf.PdfDocument;
import com.spire.pdf.PdfPaddings;
import com.spire.pdf.graphics.PdfBrushes;
import com.spire.pdf.graphics.PdfTrueTypeFont;
import com.spire.pdf.grid.PdfGrid;
import com.spire.pdf.grid.PdfGridRow;
import com.spire.pdf.grid.PdfGridRowStyle;
import com.spire.pdf.grid.PdfGridStyleBase;

import java.awt.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * PDF文件输出器
 * 
 * @since 2022年5月23日
 * @author 赵凡
 * @version 1.0
 *
 */
public class PdfWriter {

	// PDF文档对象
	private PdfDocument doc;

	/**
	 * 创建PDF文件输出对象
	 *
	 */
	protected PdfWriter() {
		doc = new PdfDocument();

		// 添加空白页（此页有水印，输出PDF时删除该页）
		doc.getPages().add();
		// 添加第一页
		doc.getPages().add();
	}


	/**
	 * 创建PDF文件输出对象
	 *
	 * @return PDF文件输出对象
	 */
	public static PdfWriter getInstance() {
		return new PdfWriter();
	}

	
	/**
	 * 将内存中的PDF对象写入指定的文件中
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
	 * 将内存中的PDF对象写入输出流中
	 * 
	 * @param out
	 *            输出流
	 */
	public void writeStream(OutputStream out) {
		try {
			doc.getPages().removeAt(0); // 删除第一个水印页
			if (out != null) {
				doc.saveToStream(out);
			}
		} finally {
			doc.dispose();
		}
	}

	
	// ======================  表格相关操作    start  =====================
	
	public <T> PdfWriter writeTable(String[] titles, PdfGridStyleBase headersStyle, List<T> rows, RowMapper<T> rowMapper, PdfGridRowStyle dataStyle) {
		PdfGrid grid = new PdfGrid();
		// 设置重复表头（表格跨页时）
		grid.setRepeatHeader(true);
		grid.getStyle().setCellPadding(new PdfPaddings(1,1,1,1));

		grid.getColumns().add(titles.length);
		PdfGridRow headRow = grid.getHeaders().add(1)[0];
		for (int i = 0; i < titles.length; i++) {
			headRow.getCells().get(i).setValue(titles[i]);
		}

		if (headersStyle == null) {
			headersStyle = new PdfGridRowStyle();
			headersStyle.setFont(new PdfTrueTypeFont(new Font("宋体", Font.BOLD,10), true));
			headersStyle.setTextBrush(PdfBrushes.getBlack());
		}
		grid.getHeaders().applyStyle(headersStyle);

		if (dataStyle == null) {
			dataStyle = new PdfGridRowStyle();
			dataStyle.setFont(new PdfTrueTypeFont(new Font("宋体", Font.PLAIN,8), true));
			dataStyle.setTextBrush(PdfBrushes.getBlack());
		}

		// 添加数据到表格
		for (int rowNum = 0; rowNum < rows.size(); rowNum++) {
			PdfGridRow r = grid.getRows().add();
			r.setStyle(dataStyle);
			String[] cols = rowMapper.map(rows.get(rowNum));
			for (int colNum = 0; colNum < cols.length; colNum++) {
				r.getCells().get(colNum).setValue(cols[colNum]);
			}
		}

		// 在PDF页面绘制表格
		grid.draw(doc.getPages().get(doc.getPages().getCount() - 1), 0, 40);

		return this;
	}

	public <T> PdfWriter writeTable(String[] titles, List<T> rows, RowMapper<T> rowMapper) {
		return writeTable(titles, null, rows, rowMapper, null);
	}

	// ======================  表格相关操作     end  =====================

	public static void main(String[] args) {
		PdfWriter w = getInstance();
		String[] titles = new String[] {"序号", "姓名", "学员状态", "缴费状态", "标准", "优惠", "实收", "支出", "退费", "待收", "代付", "剩余", "证件号", "联系电话", "申领类型-班型", "校区", "报名时间", "报名点", "招生人", "招生来源", "学员分组", "学员备注"};
		List<String[]> rows = new ArrayList<>();
		for (int i = 0; i < 60; i++) {
			String[] row = new String[titles.length];
			row[0] = (i + 1) + "";
			row[1] = "二百二十三";
			row[2] = "预报名";
			row[3] = "无单据";
			row[4] = 12016 + "";
			row[5] = 2016 + "";
			row[6] = 12016 + "";
			row[7] = 12016 + "";
			row[8] = 16 + "";
			row[9] = 16 + "";
			row[10] = 16 + "";
			row[11] = 16 + "";
			row[12] = "77*****17";
			row[13] = "999*****619";
			row[14] = "从业资格-客运-从业资格周末班";
			row[15] = "总校";
			row[16] = "2022/6/5";
			rows.add(row);
		}
		w.writeTable(titles, rows, rowObject -> rowObject);
		w.writeFile(new File("C:\\Users\\MC\\Desktop\\output.pdf"));
	}
	
}
