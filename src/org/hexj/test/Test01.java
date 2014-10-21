package org.hexj.test;

import org.hexj.excelhandler.reader.Excel2007MergeCellReaderHandler;

public class Test01 {

	public static void main(String[] args) {
		System.out.println("A1".compareTo("B2"));
		System.out.println("C3".compareTo("B2"));
		System.out.println("A3".compareTo("B2"));
		Excel2007MergeCellReaderHandler handler=new Excel2007MergeCellReaderHandler();
		try {
			handler.process("/Users/Eric/Downloads/test07.xlsx");
		} catch (Exception e) {
			e.printStackTrace();
		}
		Object obj=handler.getMergeList();
		System.out.println(handler.getMergeList());
	}
}
