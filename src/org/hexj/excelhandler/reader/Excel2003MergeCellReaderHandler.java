package org.hexj.excelhandler.reader;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.poi.hssf.eventusermodel.HSSFListener;
import org.apache.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.poi.hssf.record.BOFRecord;
import org.apache.poi.hssf.record.BoundSheetRecord;
import org.apache.poi.hssf.record.MergeCellsRecord;
import org.apache.poi.hssf.record.Record;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 抽象Excel2003读取器，通过实现HSSFListener监听器，采用事件驱动模式解析excel2003
 * 中的内容，遇到特定事件才会触发，大大减少了内存的使用。
 * 
 */
public class Excel2003MergeCellReaderHandler implements HSSFListener {
	private Map<Integer,List<CellRangeAddress>> mergeList=new HashMap<Integer,List<CellRangeAddress>>();
	
	private POIFSFileSystem fs;

	/** Should we output the formula, or the value it has? */
	private boolean outputFormulaValues = true;

	/** For parsing Formulas */
	private SheetRecordCollectingListener workbookBuildingListener;
	// excel2003工作薄
	private HSSFWorkbook stubWorkbook;

	private FormatTrackingHSSFListener formatListener;

	// 表索引
	private int sheetIndex = -1;
	private BoundSheetRecord[] orderedBSRs;
	private ArrayList boundSheetRecords = new ArrayList();

	/**
	 * 遍历excel下所有的sheet
	 * 
	 * @throws IOException
	 */
	public void process(String fileName) throws IOException {
		this.fs = new POIFSFileSystem(new FileInputStream(fileName));
		MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(
				this);
		formatListener = new FormatTrackingHSSFListener(listener);
		HSSFEventFactory factory = new HSSFEventFactory();
		HSSFRequest request = new HSSFRequest();
		if (outputFormulaValues) {
			request.addListenerForAllRecords(formatListener);
		} else {
			workbookBuildingListener = new SheetRecordCollectingListener(
					formatListener);
			request.addListenerForAllRecords(workbookBuildingListener);
		}
		factory.processWorkbookEvents(request, fs);
	}

	/**
	 * HSSFListener 监听方法，处理 Record
	 */
	@SuppressWarnings("unchecked")
	public void processRecord(Record record) {
		switch (record.getSid()) {
		case BoundSheetRecord.sid:
			boundSheetRecords.add(record);
			break;
		case BOFRecord.sid:
			BOFRecord br = (BOFRecord) record;
			if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
				// 如果有需要，则建立子工作薄
				if (workbookBuildingListener != null && stubWorkbook == null) {
					stubWorkbook = workbookBuildingListener
							.getStubHSSFWorkbook();
				}
				sheetIndex++;
				mergeList.put(sheetIndex, new ArrayList<CellRangeAddress>());
				if (orderedBSRs == null) {
					orderedBSRs = BoundSheetRecord
							.orderByBofPosition(boundSheetRecords);
				}
			}
			break;

		case MergeCellsRecord.sid:
			MergeCellsRecord mergeCellsRecord = (MergeCellsRecord) record;
			for (int i = 0; i < mergeCellsRecord.getNumAreas(); i++) {
				CellRangeAddress cellAddr = mergeCellsRecord.getAreaAt(i);
				this.mergeList.get(sheetIndex).add(cellAddr);
			}
			break;
		default:
			break;
		}
		
	}

	public Map<Integer, List<CellRangeAddress>> getMergeList() {
		return mergeList;
	}
}