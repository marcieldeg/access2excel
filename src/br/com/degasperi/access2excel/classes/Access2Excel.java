package br.com.degasperi.access2excel.classes;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import com.healthmarketscience.jackcess.Column;
import com.healthmarketscience.jackcess.Database;
import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.Table;

public class Access2Excel implements AutoCloseable {
	private Database database;
	private OutputFormat outputFormat = OutputFormat.XLSX;
	private StepWriter writer = new StepWriter(){};

	public OutputFormat getOutputFormat() {
		return outputFormat;
	}

	public void setOutputFormat(OutputFormat outputFormat) {
		this.outputFormat = outputFormat;
	}

	public void setWriter(StepWriter writer) {
		this.writer = writer;
	}

	public void open(File inputFile) throws IOException {
		database = DatabaseBuilder.open(inputFile);
	}

	public void open(String inputFile) throws IOException {
		open(new File(inputFile));
	}

	public void convert(File outputFile) throws IOException {
		try (Workbook workbook = OutputFormat.XLSX.equals(outputFormat) ? new SXSSFWorkbook() : new HSSFWorkbook()) {
			// header style
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBold(true);
			style.setFont(font);

			// date cell style
			CellStyle dateCellStyle = workbook.createCellStyle();
			short df = workbook.createDataFormat().getFormat(new SimpleDateFormat().toLocalizedPattern());
			dateCellStyle.setDataFormat(df);

			for (String tableName : database.getTableNames()) {
				writer.write(String.format("Table %s...", tableName));
				Table table = database.getTable(tableName);
				Sheet sheet = workbook.createSheet(tableName);

				// headers
				org.apache.poi.ss.usermodel.Row header = sheet.createRow(0);
				table.getColumns().forEach(o -> {
					Cell cell = header.createCell(o.getColumnIndex());
					cell.setCellValue(o.getName());
					cell.setCellStyle(style);
				});

				// data
				for (int i = 0; i < table.getRowCount(); i++) {
					com.healthmarketscience.jackcess.Row accessRow = table.getNextRow();
					org.apache.poi.ss.usermodel.Row excelRow = sheet.createRow(i + 1);

					for (int j = 0; j < accessRow.size(); j++) {
						Cell cell = excelRow.createCell(j);
						Column column = table.getColumns().get(j);
						switch (column.getType()) {
						case BOOLEAN:
							Boolean booleanValue = accessRow.getBoolean(column.getName());
							if (booleanValue != null)
								cell.setCellValue(booleanValue);
							break;
						case BYTE:
							Byte byteValue = accessRow.getByte(column.getName());
							if (byteValue != null)
								cell.setCellValue(byteValue);
							break;
						case DOUBLE:
							Double doubleValue = accessRow.getDouble(column.getName());
							if (doubleValue != null)
								cell.setCellValue(doubleValue);
							break;
						case FLOAT:
							Float floatValue = accessRow.getFloat(column.getName());
							if (floatValue != null)
								cell.setCellValue(floatValue);
							break;
						case GUID:
							cell.setCellValue(accessRow.getString(column.getName()));
							break;
						case INT:
							Short intValue = accessRow.getShort(column.getName());
							if (intValue != null)
								cell.setCellValue(intValue);
							break;
						case LONG:
							Integer longValue = accessRow.getInt(column.getName());
							if (longValue != null)
								cell.setCellValue(longValue);
							break;
						case MEMO:
							String memoValue = accessRow.getString(column.getName());
							if (memoValue != null)
								cell.setCellValue(memoValue);
							break;
						case MONEY:
							BigDecimal moneyValue = accessRow.getBigDecimal(column.getName());
							if (moneyValue != null)
								cell.setCellValue(moneyValue.doubleValue());
							break;
						case NUMERIC:
							Double numericValue = accessRow.getDouble(column.getName());
							if (numericValue != null)
								cell.setCellValue(numericValue);
							break;
						case SHORT_DATE_TIME:
							Date dateValue = accessRow.getDate(column.getName());
							if (dateValue != null)
								cell.setCellValue(dateValue);
							cell.setCellStyle(dateCellStyle);
							break;
						case TEXT:
							String textValue = accessRow.getString(column.getName());
							if (textValue != null)
								cell.setCellValue(textValue);
							break;
						case OLE:
							byte[] bytes = accessRow.getBytes(column.getName());
							if (bytes != null)
								cell.setCellValue("<" + new OleHeaderParser(bytes).getObjectName() + ">");
							break;
						default:
							cell.setCellValue("<" + column.getType().toString() + ">");
							break;
						}
					}
				}

				if (OutputFormat.XLSX.equals(outputFormat))
					((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
				table.getColumns().forEach(o -> sheet.autoSizeColumn(o.getColumnIndex()));
			}
			workbook.write(new FileOutputStream(outputFile));
			writer.write("Done.");
		} catch (Exception e) {
			writer.error(String.format("Error: %s%s", e.getClass().getSimpleName(),
					e.getMessage() == null ? "" : " - " + e.getMessage()));
		}
	}

	public void convert(String outputFile) throws IOException {
		convert(new File(outputFile));
	}

	@Override
	public void close() throws Exception {
		database.close();
	}

	public static enum OutputFormat {
		XLS, XLSX
	}

	public static interface StepWriter {
		default void write(String text) {
			System.out.println(text);
		}

		default void error(String text) {
			System.err.println(text);
		}
	}
}
