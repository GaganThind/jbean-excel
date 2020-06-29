package in.gagan.excel;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.Date;
import java.util.List;
import java.util.Set;

import in.gagan.excel.converter.Converter;
import in.gagan.excel.converter.ConverterFactory;
import in.gagan.excel.converter.ConverterType;

/**
 * The Consumer class of the API
 *
 */
public class JBeanExcel {

	public static void main(String[] args) throws ClassNotFoundException, IOException {

		JBeanExcelTest obj = new JBeanExcelTest();
		Converter convert = ConverterFactory.getConverter(ConverterType.POI_CONVERTER, obj, "<path_to_xlsx_file>");

		convert.convertBeanToExcel();
		convert.convertExcelToBean();

	}
}

class JBeanExcelTest {
	char a = 'a';
	boolean falsie = false;
	Integer[] intWrap;
	int[] intArr;
	List<String> stringList;
	List<String> stringLinkedList;
	Set<String> stringSet;
	Set<String> emptySet;
	TestClass tc;
	Integer[] intWrapEmpty = new Integer[2];
	Date dt = new Date(System.currentTimeMillis());
	Timestamp tsp = new Timestamp(System.currentTimeMillis());
	Integer intgr = Integer.MIN_VALUE;
}

class TestClass {
	char a;
	boolean falsie;
	Integer[] intWrap;
	int[] intArr;
}
