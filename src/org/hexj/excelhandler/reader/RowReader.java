package org.hexj.excelhandler.reader;

import java.util.TreeMap;

public class RowReader implements IRowReader {
	
	public RowReader() {
	}

	/*
	 * 业务逻辑实现方法
	 * 
	 * @see com.eprosun.util.excel.IRowReader#getRows(int, int, java.util.List)
	 */
	public void getRows(int sheetIndex, String sheetName, int curRow,
			TreeMap<Integer,String> rowlist) {
		System.out.println("No:"+sheetIndex+" Name:"+sheetName+" Row:"+curRow);
		for(Object obj:rowlist.keySet()){
			System.out.print("Cols"+obj.toString()+":"+rowlist.get(obj)+" ");
		}
		System.out.println();
		
	}
}