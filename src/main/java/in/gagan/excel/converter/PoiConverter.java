package in.gagan.excel.converter;

import java.io.IOException;

import in.gagan.excel.service.ExcelToJavaService;
import in.gagan.excel.service.JavaToExcelService;

/**
 * This class is used to convert the case object to excel by using apache poi
 * 
 * @author gaganthind
 *
 */
public class PoiConverter implements Converter {
	
	// Default sheet name
	private static final String DEFAULT_SHEET_NAME = "POJO";
	
	// The object for conversion
	private Object obj;

	// The path of the excel to be created
	private final String path;

	/**
	 * Default initialization constructor
	 * 
	 * @param obj
	 * @param path
	 * @throws ClassNotFoundException
	 */
	public PoiConverter(Object inputObj, final String inputPath) {
		this.obj = inputObj;
		this.path = inputPath;
	}

	/**
	 * Convert the Bean class to Excel Object
	 * @throws IOException
	 */
	@Override
	public void convertBeanToExcel() throws IOException {
		JavaToExcelService javaToExcelSvc = new JavaToExcelService(path);
		javaToExcelSvc.convertBeanFieldsToSheetRowObjs(DEFAULT_SHEET_NAME, this.obj);
		javaToExcelSvc.writeToWorkbook();
	}

	/**
	 * Convert the Excel Object to Bean Class
	 * @throws IOException
	 */
	@Override
	public void convertExcelToBean() throws IOException {
		ExcelToJavaService excelToJavaSvc = new ExcelToJavaService(path);
		excelToJavaSvc.convertSheetRowObjsToBeanFields(DEFAULT_SHEET_NAME, this.obj, null);
	}

}
