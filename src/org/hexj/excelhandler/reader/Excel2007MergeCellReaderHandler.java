package org.hexj.excelhandler.reader;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class Excel2007MergeCellReaderHandler extends DefaultHandler {
	private int sheetIndex = -1;
	private Map<Integer, List<String>> mergeList = new HashMap<Integer, List<String>>();

	/**
	 * 遍历工作簿中所有的电子表格
	 * 
	 * @param filename
	 * @throws Exception
	 */
	public void process(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader(pkg);
		XMLReader parser = fetchSheetParser();
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			sheetIndex++;
			InputStream sheet = sheets.next();
			this.mergeList.put(sheetIndex, new ArrayList<String>());
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	public XMLReader fetchSheetParser() throws SAXException {
		XMLReader parser = XMLReaderFactory
				.createXMLReader("org.apache.xerces.parsers.SAXParser");
		parser.setContentHandler(this);
		return parser;
	}

	public void startElement(String uri, String localName, String qName,
			Attributes attributes) throws SAXException {
		if ("mergeCell".equals(qName)) {
			this.mergeList.get(sheetIndex).add(attributes.getValue("ref"));
		}
	}
	public Map<Integer, List<String>> getMergeList(){
		return this.mergeList;
	}
}