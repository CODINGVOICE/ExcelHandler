package org.hexj.hdm;

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.TreeMap;
import java.util.UUID;

import org.hexj.excelhandler.reader.IRowReader;

public class RowReader implements IRowReader {
	private Connection conn;
	private int cols_count;
	private PreparedStatement stmt;
	private int n = 0;

	public RowReader(String dbUrl, String dbUser, String dbPwd,
			String insert_sql, int cols_count) {
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			this.conn = DriverManager.getConnection(dbUrl, dbUser, dbPwd);
			this.conn.setAutoCommit(false);
			this.stmt = this.conn.prepareStatement(insert_sql);
		} catch (Exception e) {
			e.printStackTrace();
			return;
		}

		this.cols_count = cols_count;
	}

	public void getRows(int sheetIndex, String sheetName, int curRow,
			TreeMap<Integer, String> rowlist) {
		if (curRow == 0)
			return;
		boolean isEpty = true;
		try {
			this.stmt.setInt(1, curRow + 1);

			this.stmt.setInt(2, sheetIndex + 1);
			this.stmt.setString(3, sheetName);
			this.stmt.setString(4, UUID.randomUUID().toString()
					.replace("-", ""));
			for (int i = 0; i < this.cols_count; i++) {
				String tmpstr = rowlist.get(i);
				if (null == tmpstr || "".equals(tmpstr.trim())) {
					this.stmt.setString(i + 5, "");
				} else {
					isEpty = false;
					this.stmt.setString(i + 5, tmpstr.trim());
				}
			}
			if (!isEpty) {
				this.stmt.addBatch();
			}
			if ((this.n + 1) % 5000 == 0) {
				this.stmt.executeBatch();
				this.conn.commit();
			}
			this.n += 1;
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void end_read() {
		try {
			this.stmt.executeBatch();
			this.conn.commit();
			this.stmt.close();
			this.conn.close();
		} catch (SQLException e) {
			e.printStackTrace();
		}
	}
}