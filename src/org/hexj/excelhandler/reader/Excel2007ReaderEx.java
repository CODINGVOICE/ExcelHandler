package org.hexj.excelhandler.reader;

import java.io.InputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class Excel2007ReaderEx extends DefaultHandler {
	// 共享字符串表
	private SharedStringsTable sst;
	// 样式表
	private StylesTable styleTable;
	// 上一次的内容
	private String lastContents;

	private int sheetIndex = -1;
	private TreeMap<Integer, String> rowlist = new TreeMap<Integer, String>();
	// 当前行
	private int curRow = 0;
	// 当前列
	private int curCol = 0;

	// 合并单元格信息
	private Map<Integer, List<String>> mergeList;

	// 当前单元格
	private String curCell;

	// 合并单元格值信息
	private Map<Integer, Map<String, String>> mergeValueList = new HashMap<Integer, Map<String, String>>();

	enum XSSFDataType {
		BOOL, ERROR, FORMULA, INLINESTR, SSTINDEX, NUMBER, DATE, NULL
	}

	private XSSFDataType nextDataType;
	private int formatIndex;
	private String formatString;
	private final DataFormatter formatter = new DataFormatter();

	private IRowReader rowReader;
	private String sheetName;

	public void setRowReader(IRowReader rowReader) {
		this.rowReader = rowReader;
	}

	/**
	 * 只遍历一个电子表格，其中sheetId为要遍历的sheet索引，从1开始，1-3
	 * 
	 * @param filename
	 * @param sheetId
	 * @throws Exception
	 */
	public void processOneSheet(String filename, int sheetId) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader(pkg);
		r.getStylesTable();
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst, styleTable);
		// 根据 rId# 或 rSheet# 查找sheet
		InputStream sheet2 = r.getSheet("rId" + sheetId);
		sheetIndex++;
		InputSource sheetSource = new InputSource(sheet2);
		parser.parse(sheetSource);
		sheet2.close();
	}

	/**
	 * 遍历工作簿中所有的电子表格
	 * 
	 * @param filename
	 * @throws Exception
	 */
	public void process(String filename) throws Exception {
		// 获取合并单元格信息
		Excel2007MergeCellReaderHandler mergeCellHandler = new Excel2007MergeCellReaderHandler();
		mergeCellHandler.process(filename);
		this.mergeList = mergeCellHandler.getMergeList();

		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader(pkg);
		StylesTable styleTable = r.getStylesTable();
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst, styleTable);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			curRow = 0;
			sheetIndex++;
			InputStream sheet = sheets.next();

			this.mergeValueList.put(sheetIndex, new HashMap<String, String>());

			InputSource sheetSource = new InputSource(sheet);
			sheetName = ((SheetIterator) sheets).getSheetName();
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst,
			StylesTable styleTable) throws SAXException {
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		this.sst = sst;
		this.styleTable = styleTable;
		parser.setContentHandler(this);
		return parser;
	}

	public void startElement(String uri, String localName, String name,
			Attributes attributes) throws SAXException {
		// c => 单元格
		if ("c".equals(name)) {
			Pattern p = Pattern.compile("[A-Z]+");
			String s = attributes.getValue("r");

			this.curCell = s;

			Matcher m = p.matcher(s);
			if (m.find()) {
				String str = m.group();
				curCol = getValue(str) - 1;
				curRow = Integer.parseInt(s.replace(str, "")) - 1;
			}
			// 默认类型
			this.nextDataType = XSSFDataType.NUMBER;
			this.formatIndex = -1;
			this.formatString = null;
			String cellType = attributes.getValue("t");
			String cellStyleStr = attributes.getValue("s");
			if ("b".equals(cellType))
				nextDataType = XSSFDataType.BOOL;
			else if ("e".equals(cellType))
				nextDataType = XSSFDataType.ERROR;
			else if ("inlineStr".equals(cellType))
				nextDataType = XSSFDataType.INLINESTR;
			else if ("s".equals(cellType))
				nextDataType = XSSFDataType.SSTINDEX;
			else if ("str".equals(cellType))
				nextDataType = XSSFDataType.FORMULA;
			if (cellStyleStr != null) {
				int styleIndex = Integer.parseInt(cellStyleStr);
				XSSFCellStyle style = this.styleTable.getStyleAt(styleIndex);
				formatIndex = style.getDataFormat();
				formatString = style.getDataFormatString();
				if (formatString == null) {
					nextDataType = XSSFDataType.NULL;
					formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
				} else if (isDateFormate(formatString)) {
					nextDataType = XSSFDataType.DATE;
					formatString = "yyyy-MM-dd HH:mm:ss";
				}
			}
		}
		// 置空
		lastContents = "";
	}
	
	private boolean isDateFormate(String formatStr){
		String formatString=formatStr.toLowerCase();
		int flag=0;
		if(formatString.contains("y")) flag++;
		if(formatString.contains("m")) flag++;
		if(formatString.contains("d")) flag++;
		if(formatString.contains("h")) flag++;
		if(formatString.contains("m")) flag++;
		if(formatString.contains("s")) flag++;
		if(flag>2) return true;
		return false;
	}

	/**
	 * 对解析出来的数据进行类型处理
	 * 
	 * @param value
	 *            单元格的值（这时候是一串数字）
	 * @param thisStr
	 *            一个空字符串
	 * @return
	 */
	public String getDataValue(String value) {
		String thisStr = "";
		switch (nextDataType) {
		case BOOL:
			char first = value.charAt(0);
			thisStr = first == '0' ? "FALSE" : "TRUE";
			break;
		case ERROR:
			thisStr = "\"ERROR:" + value.toString() + '"';
			break;
		case FORMULA:
			thisStr = '"' + value.toString() + '"';
			break;
		case INLINESTR:
			XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
			thisStr = rtsi.toString();
			rtsi = null;
			break;
		case SSTINDEX:
			String sstIndex = value.toString();
			try {
				int idx = Integer.parseInt(sstIndex);
				XSSFRichTextString rtss = new XSSFRichTextString(
						sst.getEntryAt(idx));
				thisStr = rtss.toString();
				rtss = null;
			} catch (NumberFormatException ex) {
				thisStr = value.toString();
			}
			break;
		case NUMBER:
			if (formatString != null) {
				thisStr = formatter.formatRawCellContents(
						Double.parseDouble(value), formatIndex, formatString)
						.trim();
			} else {
				thisStr = value;
			}
			thisStr = thisStr.replace("_", "").trim();
			break;
		case DATE:
			thisStr = formatter.formatRawCellContents(
					Double.parseDouble(value), formatIndex, formatString);
			break;
		default:
			thisStr = " ";
			break;
		}
		return thisStr;
	}

	public void endElement(String uri, String localName, String name)
			throws SAXException {
		if ("c".equals(name)) {
			for (String range : this.mergeList.get(this.sheetIndex)) {
				int idx = range.indexOf(':');
				String start = range.substring(0, idx);
				String end = range.substring(idx + 1);
				Pattern p = Pattern.compile("[A-Z]+");
				Matcher m = p.matcher(start);
				String colsFrom = "";
				String colsTo = "";
				String rowFrom = "";
				String rowTo = "";
				String curCols = "";
				String curRows = "";
				if (m.find()) {
					colsFrom = m.group();
					rowFrom = start.replace(colsFrom, "");
				}
				m = p.matcher(end);
				if (m.find()) {
					colsTo = m.group();
					rowTo = end.replace(colsTo, "");
				}
				m = p.matcher(curCell);
				if (m.find()) {
					curCols = m.group();
					curRows = this.curCell.replace(curCols, "");
				}
				int irowFrom=Integer.parseInt(rowFrom);
				int irowTo=Integer.parseInt(rowTo);
				int icurRows=Integer.parseInt(curRows);
				int icolsFrom=getValue(colsFrom);
				int icolsTo=getValue(colsTo);
				int icurCols=getValue(curCols);
				if (icurRows>=irowFrom
						&& icurRows<=irowTo
						&& icurCols>=icolsFrom
						&& icurCols<=icolsTo) {
					rowlist.put(curCol, this.mergeValueList
							.get(this.sheetIndex).get(range));
				}
			}
		}
		if ("v".equals(name) || "t".equals(name)) {
			String value = lastContents.trim();
			value = this.getDataValue(value);
			for (String range : this.mergeList.get(this.sheetIndex)) {
				int idx = range.indexOf(':');
				String start = range.substring(0, idx);
				if (start.equals(this.curCell)) {
					this.mergeValueList.get(this.sheetIndex).put(range, value);
				}
			}
			rowlist.put(curCol, value);
			// rowlist.put(curCol,this.getDataValue(value));
		} else if ("row".equals(name)) {
			rowReader.getRows(sheetIndex, sheetName, curRow, rowlist);
			rowlist.clear();
			curRow++;
		}

	}

	public void characters(char[] ch, int start, int length)
			throws SAXException {
		// 得到单元格内容的值
		lastContents += new String(ch, start, length);
	}

	private int getValue(String str) {
		byte[] chars = str.getBytes();
		if (chars.length > 1) {
			return getValue(str.substring(1)) + (chars[0] - 'A' + 1) * 26;
		} else {
			return chars[0] - 'A' + 1;
		}
	}
}