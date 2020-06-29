package in.gagan.excel.converter;

import java.io.IOException;

/**
 * Default interface for conversion of the bean to excel and vice-versa
 * 
 * @author gaganthind
 *
 */
public interface Converter {
	
	// Method to convert java bean object to excel
	void convertBeanToExcel() throws IOException;
	
	// Method to convert excel to java bean object
	void convertExcelToBean() throws IOException;

}
