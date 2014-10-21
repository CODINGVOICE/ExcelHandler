package org.hexj.excelhandler.reader;

public class ExcelReaderUtil {

	// excel2003扩展名
	public static final String EXCEL03_EXTENSION = ".xls";
	// excel2007扩展名
	public static final String EXCEL07_EXTENSION = ".xlsx";

	/**
	 * 读取Excel文件，可能是03也可能是07版本
	 * 
	 * @param reader 行处理实现
	 * @param fileName 文件路径
	 * @param handleMergeCell 是否展开合并单元格
	 * @throws Exception
	 */
	public static void readExcel(IRowReader reader, String fileName,
			boolean handleMergeCell) throws Exception {
		if (handleMergeCell) {
			// 处理excel2003文件
			if (fileName.endsWith(EXCEL03_EXTENSION)) {
				Excel2003ReaderEx excel03 = new Excel2003ReaderEx();
				excel03.setRowReader(reader);
				excel03.process(fileName);
				// 处理excel2007文件
			} else if (fileName.endsWith(EXCEL07_EXTENSION)) {
				Excel2007ReaderEx excel07 = new Excel2007ReaderEx();
				excel07.setRowReader(reader);
				excel07.process(fileName);
			} else {
				throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
			}
		} else {
			// 处理excel2003文件
			if (fileName.endsWith(EXCEL03_EXTENSION)) {
				Excel2003Reader excel03 = new Excel2003Reader();
				excel03.setRowReader(reader);
				excel03.process(fileName);
				// 处理excel2007文件
			} else if (fileName.endsWith(EXCEL07_EXTENSION)) {
				Excel2007Reader excel07 = new Excel2007Reader();
				excel07.setRowReader(reader);
				excel07.process(fileName);
			} else {
				throw new Exception("文件格式错误，fileName的扩展名只能是xls或xlsx。");
			}
		}
	}
}
