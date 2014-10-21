package org.hexj.excelhandler.reader;

import java.util.TreeMap;

public interface IRowReader {
	
	/**业务逻辑实现方法
	 * @param sheetIndex
	 * @param curRow
	 * @param rowlist
	 */
	public  void getRows(int sheetIndex,String sheetName,int curRow, TreeMap<Integer,String> rowlist);
}