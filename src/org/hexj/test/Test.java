package org.hexj.test;

import org.hexj.excelhandler.reader.ExcelReaderUtil;
import org.hexj.excelhandler.reader.IRowReader;
import org.hexj.excelhandler.reader.RowReader;

public class Test {

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		IRowReader reader = new RowReader();
		try {
			long start = System.currentTimeMillis();
			ExcelReaderUtil.readExcel(reader, "/Users/Eric/Downloads/人员信息.xlsx",true);
			long end = System.currentTimeMillis();
			System.out.println("Time used:"+(end-start)/1000+"s");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
