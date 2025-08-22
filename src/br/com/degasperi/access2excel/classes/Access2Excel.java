package br.com.degasperi.access2excel.classes;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

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

/**
 * Main class for converting MS Access databases to Excel spreadsheets.
 * This class handles the opening, reading, and conversion process.
 */
public class Access2Excel implements AutoCloseable {
	private Database database;
	private OutputFormat outputFormat = OutputFormat.XLSX;
	private StepWriter writer = new StepWriter(){};

	/**
	 * Gets the currently configured output format.
	 * @return The output format (XLS or XLSX).
	 */
	public OutputFormat getOutputFormat() {
		return outputFormat;
	}

	/**
	 * Sets the desired output format for the Excel file.
	 * @param outputFormat The format to use (XLS or XLSX).
	 */
	public void setOutputFormat(OutputFormat outputFormat) {
		this.outputFormat = outputFormat;
	}

	/**
	 * Sets a custom writer for logging conversion steps.
	 * @param writer The StepWriter implementation.
	 */
	public void setWriter(StepWriter writer) {
		this.writer = writer;
	}

	/**
	 * Opens an MS Access database file in read-only mode.
	 * Also performs a security check to prevent path traversal attacks.
	 *
	 * @param inputFile The Access database file to open.
	 * @throws IOException If the file cannot be opened.
	 * @throws SecurityException If the file path is considered a security risk (path traversal).
	 */
	public void open(File inputFile) throws IOException {
		String workingDir = new File(".").getCanonicalPath();
		if (!inputFile.getCanonicalPath().startsWith(workingDir))
			throw new SecurityException("Input file path is outside the working directory.");
		database = new DatabaseBuilder(inputFile).setReadOnly(true).open();
	}

	/**
	 * Convenience method to open an MS Access database from a file path string.
	 *
	 * @param inputFile The path to the Access database file.
	 * @throws IOException If the file cannot be opened.
	 * @throws SecurityException If the file path is considered a security risk (path traversal).
	 */
	public void open(String inputFile) throws IOException {
		open(new File(inputFile));
	}

	/**
	 * Converts the opened Access database to an Excel file.
	 * Iterates through all tables and writes the data to corresponding sheets.
	 *
	 * @param outputFile The file where the Excel spreadsheet will be saved.
	 * @throws IOException If there is an error during file writing.
	 * @throws SecurityException If the output file path is considered a security risk (path traversal).
	 */
	public void convert(File outputFile) throws IOException {
		String workingDir = new File(".").getCanonicalPath();
		if (!outputFile.getCanonicalPath().startsWith(workingDir))
			throw new SecurityException("Output file path is outside the working directory.");
		try (Workbook workbook = OutputFormat.XLSX.equals(outputFormat) ? new SXSSFWorkbook() : new HSSFWorkbook()) {
			// header style
			CellStyle style = workbook.createCellStyle();
			Font font = workbook.createFont();
			font.setBold(true);
			style.setFont(font);

			// date cell style
			CellStyle dateCellStyle = workbook.createCellStyle();
			short df = workbook.createDataFormat().getFormat(DateTimeFormatter.ISO_LOCAL_DATE_TIME.toString());
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
							LocalDateTime dateValue = accessRow.getLocalDateTime(column.getName());
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

	/**
	 * Convenience method to convert the database to an Excel file specified by a path string.
	 *
	 * @param outputFile The path to the output Excel file.
	 * @throws IOException If there is an error during file writing.
	 * @throws SecurityException If the output file path is considered a security risk (path traversal).
	 */
	public void convert(String outputFile) throws IOException {
		convert(new File(outputFile));
	}

	/**
	 * Closes the database connection.
	 * Should be called after conversion is complete, preferably in a try-with-resources block.
	 */
	@Override
	public void close() throws Exception {
		database.close();
	}

	/**
	 * Enum representing the supported Excel output formats.
	 */
	public static enum OutputFormat {
		XLS, XLSX
	}

	/**
	 * A simple interface for logging progress during the conversion process.
	 * Allows the caller to implement custom logging behavior.
	 */
	public static interface StepWriter {
		/**
		 * Writes a standard progress message.
		 * @param text The message to write.
		 */
		default void write(String text) {
			System.out.println(text);
		}

		/**
		 * Writes an error message.
		 * @param text The error message to write.
		 */
		default void error(String text) {
			System.err.println(text);
		}
	}
}
