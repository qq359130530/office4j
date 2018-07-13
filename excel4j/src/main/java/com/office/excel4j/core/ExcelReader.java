package com.office.excel4j.core;

import java.io.IOException;
import java.io.InputStream;
import java.io.PushbackInputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.LabelSSTRecord;
import org.apache.poi.hssf.record.NumberRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.record.SSTRecord;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.NPOIFSFileSystem;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import com.office.excel4j.core.exception.AccessCellException;
import com.office.excel4j.core.exception.AccessExcelException;

/**
 * Excel文件读取器
 * 
 * @since 2018年7月3日
 * @author 赵凡
 * @version 1.0
 *
 */
public abstract class ExcelReader extends DefaultHandler implements HSSFListener {

	private SSTRecord sstrec;// XLS版共享字符串池

	private List<String> sheetNames = new ArrayList<>();// Sheet名称
	
	private Integer curSheetIdx = -1;// 当前Sheet索引，从0开始，-1表示没有sheet
	
	private AccessExcelException accessExcelException;// Excel读取异常
	
	private boolean fastFailure = false;// 遇到异常是否快速失败
	
	private SharedStringsTable sst;// XLSX版共享字符串池
	
	private String lastContents;
	
	private boolean nextIsString;
	
	/**
	 * 读取Excel输入流
	 * 
	 * @param input
	 *            Excel输入流
	 * @throws AccessExcelException 
	 */
	public void read(InputStream input) throws AccessExcelException {
		try {
			// If clearly doesn't do mark/reset, wrap up
	        if (! input.markSupported()) {
	        	input = new PushbackInputStream(input, 8);
	        }

	        // Ensure that there is at least some data there
	        byte[] header8 = IOUtils.peekFirst8Bytes(input);

	        // Try to create
	        if (NPOIFSFileSystem.hasPOIFSHeader(header8)) {
	        	try (
	    				POIFSFileSystem poifs = new POIFSFileSystem(input);// 初始化POI文件系统
	    				InputStream din = poifs.createDocumentInputStream("Workbook")// 创建Excel工作簿输入流
	    			) {
	    			// 创建读请求
	    	        HSSFRequest req = new HSSFRequest();
	    	        // 监听所有记录读取事件
	    	        req.addListenerForAllRecords(this);
	    	        // 创建事件工厂
	    	        HSSFEventFactory factory = new HSSFEventFactory();
	    	        // 基于文档流处理事件
	    	        factory.processEvents(req, din);
	    		} catch (IOException e) {
	    			throw new AccessExcelException("读取Excel失败！", e);
	    		}
	        } else if (DocumentFactoryHelper.hasOOXMLHeader(input)) {
	        	try (OPCPackage pkg = OPCPackage.open(input);) {
					XSSFReader r = new XSSFReader(pkg);
					sst = r.getSharedStringsTable();
					XMLReader parser = fetchSheetParser(sst);
					Iterator<InputStream> sheets = r.getSheetsData();
					while(sheets.hasNext()) {
						InputStream sheet = sheets.next();
						InputSource sheetSource = new InputSource(sheet);
						try {
							parser.parse(sheetSource);
						} catch (Exception e1) {
							throw new AccessExcelException("解析Sheet失败！", e1);
						} finally {
							if (sheet != null) {
								try {
									sheet.close();
								} catch (IOException e1) {
									
								}
							}
						}
					}
				} catch (Exception e1) {
					throw new AccessExcelException("读取Excel失败：无法识别的文件格式！", e1);
				} 
	        } else {
	        	throw new AccessExcelException("不能失败的Excel文件格式！");
	        }
		} catch (Exception e) {
			throw new AccessExcelException("读取Excel失败！", e);
		}
	}
	
	// 创建SAX XMLReader对象
	private XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException {
		XMLReader parser =
			XMLReaderFactory.createXMLReader(
					"org.apache.xerces.parsers.SAXParser"
			);
		parser.setContentHandler(this);
		return parser;
	}
	
	// XML元素开始
	public final void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		// c => cell
		if(name.equals("c")) {
			// Print the cell reference
			System.out.print(attributes.getValue("r") + " - ");
			// Figure out if the value is an index in the SST
			String cellType = attributes.getValue("t");
			if(cellType != null && cellType.equals("s")) {
				nextIsString = true;
			} else {
				nextIsString = false;
			}
		}
		// Clear contents cache
		lastContents = "";
	}
	
	// XML元素结束
	public final void endElement(String uri, String localName, String name)
			throws SAXException {
		// Process the last contents as required.
		// Do now, as characters() may be called more than once
		if(nextIsString) {
			int idx = Integer.parseInt(lastContents);
			lastContents = new XSSFRichTextString(sst.getEntryAt(idx)).toString();
			nextIsString = false;
		}

		// v => contents of a cell
		// Output after we've seen the string contents
		if(name.equals("v")) {
			System.out.println(lastContents);
		}
	}

	// XML元素内容节点
	public final void characters(char[] ch, int start, int length)
			throws SAXException {
		lastContents += new String(ch, start, length);
	}
	
	/**
	 * 遇到异常是否快速失败
	 * 
	 * @return 遇到异常是否快速失败
	 */
	public boolean isFastFailure() {
		return fastFailure;
	}

	/**
	 * 设置遇到异常是否快速失败
	 * 
	 * @param fastFailure
	 *            遇到异常是否快速失败
	 */
	public void setFastFailure(boolean fastFailure) {
		this.fastFailure = fastFailure;
	}

	/**
	 * 获取Sheet索引，从1开始
	 * 
	 * @return Sheet索引，从1开始
	 */
	public int getSheetIdx() {
		return curSheetIdx + 1;
	}
	
	/**
	 * 获取Sheet名称
	 * 
	 * @return Sheet名称
	 */
	public String getSheetName() {
		return sheetNames.get(curSheetIdx);
	}
	
	/**
	 * 添加异常
	 * 
	 * @param cell
	 *            单元格
	 * @param message
	 *            异常消息
	 * @return 异常读取器
	 */
	protected ExcelReader addException(Cell cell, String message) {
		accessExcelException.addCellException(cell, message);
		return this;
	}
	
	@Override
	public final void processRecord(Record record) {
		switch (record.getSid()) {
			case BOFRecord.sid:
	            BOFRecord bof = (BOFRecord) record;
	            if (bof.getType() == BOFRecord.TYPE_WORKSHEET) {// Sheet开始位置
	                curSheetIdx ++;
	            }
	            break;
            case BoundSheetRecord.sid:
            	// Sheet边界记录
                BoundSheetRecord bsr = (BoundSheetRecord) record;
                sheetNames.add(bsr.getSheetname());
                break;
            case NumberRecord.sid:
            	// 数值记录
                NumberRecord numrec = (NumberRecord) record;
                tryProcessCell(new Cell(getSheetIdx(), getSheetName(), numrec.getRow(), numrec.getColumn(), numrec.getValue()));
                break;
            case LabelSSTRecord.sid:
            	// 字符串记录
                LabelSSTRecord lrec = (LabelSSTRecord) record;
                tryProcessCell(new Cell(getSheetIdx(), getSheetName(), lrec.getRow(), lrec.getColumn(), sstrec.getString(lrec.getSSTIndex())));
                break;
            case SSTRecord.sid:
            	// 共享字符串池
            	sstrec = (SSTRecord) record;
                break;
        }
	}
	
	/**
	 * 尝试处理单元格
	 * 
	 * @param cell
	 *            单元格
	 */
	protected final void tryProcessCell(Cell cell) {
		try {
			processCell(cell);
		} catch (Exception e) {
			if (fastFailure) {// 遇到异常快速失败
				throw e;
			} 
			
			// 非快速失败模式下统计所有遇到的异常
			if (e instanceof AccessCellException) {
				accessExcelException.addCellException((AccessCellException) e);
			} else {
				accessExcelException.addCellException(cell, e.getMessage());
			}
		}
	}
	
	/**
	 * 处理单元格
	 * 
	 * @param cell
	 *            单元格
	 * @throws AccessCellException
	 *             访问单元格异常
	 */
	protected abstract void processCell(Cell cell) throws AccessCellException;
	
}