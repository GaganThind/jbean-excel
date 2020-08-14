# jbean-excel
This project is aimed to convert the java bean fields with data to excel and then reconvert them back to java bean. The project would be able to convert the
below mentioned datastructures and fields present inside a java object to corresponding excel file.
(1) Array
(2) Enum
(3) List
(4) Set
(5) Class: Java and User Defined
(6) Primitives and Wrappers

In Progresss:
(1) Map
(2) Other Types

For Complex objects which can be further divided into sub objects i.e. User Defined class, Set, List and Arrays. 
The API will create different sheet for the object type and each Collection or Array object will be represented. 
If we have multiple objects (in case of collections), then a ~ (tilda) row would denote multiple instances inside the collection.

Usage:

Converter convert = ConverterFactory.getConverter(ConverterType.POI_CONVERTER, <Java Object>, <PATH_TO_EXCEL>);

This factory method takes 3 parameters
(i)   ConverterType: Type of the converter to be used
(ii)  <Java Object>: Java Object to convert to excel
(iii) <PATH_TO_EXCEL>: Provide the path of the excel. This is where to save/retreive the excel from. 

This methid will return the implementation class instance of the converter interface which provide 2 basic methods

1) Convert java bean to excel:
  (a) convert.convertBeanToExcel();
  This Method simply converts the provided jaba object to the excel workbook.
  
2) Convert excel to java bean:
  (a) convert.convertExcelToBean();
  This Method will read the provided excel and add the data to the provided object.

