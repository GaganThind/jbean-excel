package in.gagan.excel.service;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Abstract service class with common methods
 * 
 * @author gaganthind
 *
 */
public abstract class AbstractService {

	/**
	 * Common constants
	 * */
	protected static final String SKIP_LIST_SHEET_NAME = "Skip List";
	protected static final String FIELDS_CANNOT_READ_SHEET_NAME = "Fields Not Readable";
	protected static final String FIELDS_CANNOT_WRITE_SHEET_NAME = "Fields Not Writable";
	protected static final String FIELDS_NOT_FOUND_SHEET_NAME = "Fields Not Found";
	protected static final String COLUMN_HEADER_NAME = "Name";
	protected static final String COLUMN_HEADER_TYPE = "Type";
	protected static final String COLUMN_HEADER_VALUE = "Value";
	protected static final String COLUMN_VALUE_TILDA = "~";
	
	// Default date/Timestamp saving format
	public final DateFormat dateFormat = new SimpleDateFormat("MM-dd-yyyy hh:mm:ss.SSS");

	// This skip list maintains the fields that are not to be stored in excel
	protected final List<String> skipList = Arrays.asList("serialVersionUID");

	// Default Workbook
	protected Workbook excelWorkbook;

	// Workbook path to save and load
	protected final String path;

	/**
	 * Param Constructor
	 * 
	 * @param path
	 */
	protected AbstractService(final String path) {
		this.path = path;
	}

	/**
	 * Write data to Excel workbook
	 *
	 * @throws IOException
	 */
	public void writeToWorkbook() throws IOException {
		try (FileOutputStream outputStream = new FileOutputStream(path)) {
			excelWorkbook.write(outputStream);
		} catch (IOException e) {
			throw new IOException(" Unable to write data to excel on path: " + path + " : " + e.getMessage());
		} finally {
			excelWorkbook.close();
		}
	}

}
