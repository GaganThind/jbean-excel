package in.gagan.excel.service;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.sql.Timestamp;
import java.util.Collection;
import java.util.Date;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.ClassUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import in.gagan.excel.util.HelperUtils;

/**
 * Service class to convert java beans to excel
 * 
 * @author gaganthind
 *
 */
public class JavaToExcelService extends AbstractService {

	/**
	 * Param constructor
	 * 
	 * @throws IOException
	 */
	public JavaToExcelService(final String path) throws IOException {
		super(path);
		excelWorkbook = WorkbookFactory.create(true);
	}

	/**
	 * Convert the fields to rows in excel object
	 *
	 * @param sheetName
	 * @param javaObj
	 */
	public void convertBeanFieldsToSheetRowObjs(String sheetName, Object javaObj) {

		int rowCount = 0;
		Sheet sheet = createOrGetSheet(sheetName);

		// Local variables
		String fieldName = null;
		Object fieldValue = null;
		Class<?> fieldType = null;
		Field[] fields = null != javaObj.getClass() ? javaObj.getClass().getDeclaredFields() : new Field[0];

		for (Field field : fields) {

			// If static and final then don't include
			fieldName = field.getName();

			// Add fields to skipList to not convert to excel
			if (skipList.contains(fieldName) || HelperUtils.isStaticFinalField(field)) {
				continue;
			}

			try {
				fieldValue = FieldUtils.readField(field, javaObj, true);
			} catch (IllegalAccessException e) {
				continue;
			}

			// For interfaces, set the implementation class
			fieldType = getFieldType(field, fieldValue);

			if (null != fieldValue) {
				
				if (ClassUtils.isAssignable(fieldType, Date.class) || ClassUtils.isAssignable(fieldType, Timestamp.class)) {
					fieldValue = dateFormat.format(fieldValue);
				} else if (ClassUtils.isAssignable(fieldType, Collection.class)) {
					fieldValue = handleCollection(field, fieldType, fieldValue, fieldName);
				} else if (fieldType.isArray()) {
					fieldValue = handleArray(fieldType, fieldValue, fieldName);
				} else if (HelperUtils.isComplexObject(fieldType)) {
					fieldValue = handleComplexObject(fieldType, fieldValue, fieldName);
				}
				
			}

			// Create a row for each bean field
			createSheetRowFromFieldData(sheet, ++rowCount, String.valueOf(fieldType), String.valueOf(fieldName), String.valueOf(fieldValue));
		}
		
	}
	
	/**
	 * Create a new sheet or return the existing sheet
	 * 
	 * @param sheetName
	 * @return
	 */
	private Sheet createOrGetSheet(String sheetName) {
		Sheet sheet;
		int rowCount = 0;
		
		if (null == excelWorkbook.getSheet(sheetName)) {
			sheet = excelWorkbook.createSheet(sheetName);

			// Create Header row for new sheet
			createSheetRowFromFieldData(sheet, rowCount, COLUMN_HEADER_TYPE, COLUMN_HEADER_NAME, COLUMN_HEADER_VALUE);

		} else {
			sheet = excelWorkbook.getSheet(sheetName);
			rowCount = sheet.getLastRowNum() + 1;

			// If the sheet already exists then add tilda row to separate the different
			// objects
			createSheetRowFromFieldData(sheet, rowCount, COLUMN_VALUE_TILDA, COLUMN_VALUE_TILDA, COLUMN_VALUE_TILDA);
		}
		
		return sheet;
	}

	/**
	 * Create a excel row with cell data
	 *
	 * @param sheet
	 * @param rowNum
	 * @param type
	 * @param name
	 * @param value
	 */
	private void createSheetRowFromFieldData(Sheet sheet, int rowNum, String type, String name, String value) {

		Row row = sheet.createRow(rowNum);

		// Add data to first column for field type
		Cell cell = row.createCell(0);
		cell.setCellValue(type);

		// Add data to second column for field name
		if (null != name) {
			cell = row.createCell(1);
			cell.setCellValue(name);
		}

		// Add data to third column for field data
		if (null != value) {
			cell = row.createCell(2);
			cell.setCellValue(value);
		}
	}

	/**
	 * Convert array object to row
	 *
	 * @param sheetName
	 * @param collectionType
	 * @param colObj
	 */
	private void convertArrayOrCollectionToRowObjs(String sheetName, Object object) {
		
		if (object instanceof Object[]) {
			for (Object obj : (Object[]) object) {
				convertBeanFieldsToSheetRowObjs(sheetName, obj);
			}
		} else if (object instanceof Collection) {
			for (Object coll : (Collection<?>) object) {
				convertBeanFieldsToSheetRowObjs(sheetName, coll);
			}
		}
		
	}
	
	/**
	 * For interfaces, set the implementation class
	 * 
	 * @param field
	 * @param fieldValue
	 * @return
	 */
	private Class<?> getFieldType(Field field, Object fieldValue) {
		return (!field.getType().isInterface() || null == fieldValue) ? field.getType() : fieldValue.getClass();
	}
	
	/**
	 * Handle the logic to loop over arrays or collections
	 * 
	 * @param object
	 * @param fieldName
	 * @return
	 */
	private String handleArrayOrCollectionObjects(Object object, String fieldName) {
		
		// Due to excel sheet name length truncate it to max 30 chars
		String sheetNameTrunc = fieldName.length() > 30 ? fieldName.substring(0, 30) : fieldName;
		convertArrayOrCollectionToRowObjs(sheetNameTrunc, object);

		// The value of this field should be the name of the sheet
		return new StringBuilder("Sheet:").append(sheetNameTrunc).toString();
	}
	
	/**
	 * Handle Array
	 * 
	 * @param fieldType
	 * @param fieldValue
	 * @param fieldName
	 * @return
	 */
	private Object handleArray(Class<?> fieldType, Object fieldValue, String fieldName) {
		
		Object value = fieldValue;
		
		if (HelperUtils.isComplexObject(fieldType.getComponentType())) {
			Object[] objArr = (Object[]) ConvertUtils.convert(fieldValue, fieldType);

			if (objArr.length > 0) {
				value = handleArrayOrCollectionObjects(objArr, fieldName);
			}
		} else {
			value = (ClassUtils.isPrimitiveOrWrapper(fieldType.getComponentType()) || fieldType.getComponentType().isEnum())
							? HelperUtils.writeArrayObjectAsString(fieldValue)
							: fieldValue.toString();
		}
		
		return value;
		
	}
	
	/**
	 * Handle collections
	 * 
	 * @param field
	 * @param fieldType
	 * @param fieldValue
	 * @param fieldName
	 * @return
	 */
	private Object handleCollection(Field field, Class<?> fieldType, Object fieldValue, String fieldName) {
		
		Object value = fieldValue;
		
		Class<?> genericType = (Class<?>) ((ParameterizedType) field.getGenericType()).getActualTypeArguments()[0];

		// If generic type is primitive, enum or String, then don't create a new sheet.
		// New sheet should be created for complex objects only.
		if (HelperUtils.isComplexObject(genericType)) {
			Collection<?> collection = (Collection<?>) ConvertUtils.convert(fieldValue, fieldType);
			
			if (!collection.isEmpty()) {
				value = handleArrayOrCollectionObjects(collection, fieldName);
			}
		}
		
		return value;
	}
	
	/**
	 * Handle the logic to display data of complex objects like a custom class
	 * 
	 * @param fieldType
	 * @param fieldValue
	 * @param fieldName
	 * @return
	 */
	private Object handleComplexObject(Class<?> fieldType, Object fieldValue, String fieldName) {
		Object object = ConvertUtils.convert(fieldValue, fieldType);
		
		// Due to excel sheet name length truncate it to max 30 chars
		String sheetNameTrunc = fieldName.length() > 30 ? fieldName.substring(0, 30) : fieldName;
		convertBeanFieldsToSheetRowObjs(sheetNameTrunc, object);
		
		// The value of this field should be the name of the sheet
		return new StringBuilder("Sheet:").append(sheetNameTrunc).toString();
	}


}
