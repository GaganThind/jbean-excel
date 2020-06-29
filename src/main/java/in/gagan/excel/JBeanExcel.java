package in.gagan.excel;

import java.io.IOException;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.LinkedList;
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
		Converter convert = ConverterFactory.getConverter(ConverterType.POI_CONVERTER, obj, "Test.xlsx");

		convert.convertBeanToExcel();
		convert.convertExcelToBean();

	}
}

class JBeanExcelTest {
	char a = 'a';
	boolean falsie = false;
	Integer[] intWrap = new Integer[] { 1, 2, 3, 4, 5 };
	int[] intArr = new int[] { 6, 7, 8, 9, 10 };
	List<String> stringList = new ArrayList<>(Arrays.asList("11", "12", "13"));
	List<String> stringLinkedList = new LinkedList<>(Arrays.asList("14", "15", "16"));
	Set<String> stringSet = new HashSet<>(Arrays.asList("17", "18", "19"));
	Set<String> emptySet;
	TestClass tc = new TestClass();
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
