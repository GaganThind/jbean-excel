package in.gagan.excel.service;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.ParameterizedType;
import java.sql.Timestamp;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.ClassUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.reflect.ConstructorUtils;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import in.gagan.excel.util.HelperUtils;

/**
 * Service class to converts the excel to java bean object
 * 
 * @author gaganthind
 *
 */
public class ExcelToJavaService extends AbstractService {

	/**
	 * Param constructor
	 * 
	 * @param path
	 * @throws IOException
	 */
	public ExcelToJavaService(final String path) throws IOException {
		super(path);
		excelWorkbook = WorkbookFactory.create(new File(path));
	}

	/**
	 * Convert the excel sheet rows to object fields
	 *
	 * @param sheetName
	 * @param javaObj
	 * @param subType
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void convertSheetRowObjsToBeanFields(String sheetName, Object javaObj, Class<?> subType) {

		Object baseObj = null;
		Sheet sheet = excelWorkbook.getSheet(sheetName);

		if (javaObj instanceof Collection) {
			try {
				baseObj = createInstance(subType);
				((Collection) javaObj).add(baseObj);
			} catch (InstantiationException | IllegalAccessException | NoSuchMethodException | SecurityException | 
					IllegalArgumentException | InvocationTargetException e) {
				return;
			}
		} else {
			baseObj = javaObj;
		}

		// Local variables
		Row row = null;
		Field field = null;
		String fieldValue = null;
		Class<?> fieldType = null;
		String fieldName = null;
		String fieldTypeName = null;

		// Iterate and convert to bean
		Iterator<Row> rowIterator = sheet.rowIterator();

		while (rowIterator.hasNext()) {

			row = rowIterator.next();

			// Skip first row containing headers
			if (0 == row.getRowNum()) {
				continue;
			}

			fieldName = getValueFromCellasString(row.getCell(1));

			// New Object is created when a tilda is encountered.
			if (COLUMN_VALUE_TILDA.equals(fieldName)) {
				try {
					baseObj = createInstance(subType);
					((Collection) javaObj).add(baseObj);
				} catch (InstantiationException | IllegalAccessException | NoSuchMethodException | SecurityException | 
						IllegalArgumentException | InvocationTargetException e) {
					return;
				}
				continue;
			}
			
			field = FieldUtils.getDeclaredField(baseObj.getClass(), fieldName, true);
			
			if (null == field) {
				continue;
			}

			// Get exact value from cell
			fieldValue = getValueFromCellasString(row.getCell(2));
			
			// Name of the Field
			fieldTypeName = getValueFromCellasString(row.getCell(0));

			try {
				// For interfaces, set the implementation class
				fieldType = (!field.getType().isInterface() || null == fieldValue) ? field.getType()
						: Class.forName(fieldTypeName.split(" ")[1]);

				setRowValueInBeanField(baseObj, field, fieldValue, fieldType);

			} catch (IllegalArgumentException | IllegalAccessException | ParseException | InstantiationException
					| ClassNotFoundException | NoSuchMethodException | InvocationTargetException e) {
				// Do nothing
			}

		}
	}

	/**
	 * Set the value from excel to object
	 *
	 * @param obj
	 * @param field
	 * @param fieldValue
	 * @param fieldType
	 * @throws IllegalArgumentException
	 * @throws IllegalAccessException
	 * @throws ParseException
	 * @throws InvocationTargetException 
	 * @throws NoSuchMethodException 
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void setRowValueInBeanField(Object obj, Field field, String fieldValue, Class<?> fieldType)
			throws IllegalAccessException, ParseException, InstantiationException, NoSuchMethodException, InvocationTargetException {

		if (null == fieldValue) {
			field.set(obj, null);
		} else if (ClassUtils.isAssignable(fieldType, Date.class)) {
			field.set(obj, dateFormat.parse(fieldValue));
		} else if (ClassUtils.isAssignable(fieldType, Timestamp.class)) {
			field.set(obj, new Timestamp(dateFormat.parse(fieldValue).getTime()));
		} else if (fieldType.isEnum()) {
			field.set(obj, Enum.valueOf((Class<Enum>) fieldType, fieldValue));
		} else if (ClassUtils.isAssignable(fieldType, Collection.class)) {
			Collection<?> coll = handleCollection(field, fieldType, fieldValue);
			field.set(obj, coll);
		} else if (fieldType.isArray()) {
			Object objArr = handleArray(fieldType, fieldValue);
			field.set(obj, objArr);
		} else if (HelperUtils.isComplexObject(fieldType)) {
			Object objectInstance = createInstance(fieldType);
			String sheetName = fieldValue.split(":")[1];
			convertSheetRowObjsToBeanFields(sheetName, objectInstance, null);
			field.set(obj, ConvertUtils.convert(objectInstance, fieldType));
		} else {

			// Primitive and java classes
			field.set(obj, ConvertUtils.convert(fieldValue, fieldType));
		}
		
	}

	/**
	 * Return value based on cell type
	 *
	 * @param cell
	 * @return
	 */
	private Object getValueFromCell(Cell cell) {

		Object value = null;

		switch (cell.getCellType()) {
		case BOOLEAN:
			value = cell.getBooleanCellValue();
			break;

		case NUMERIC:
			value = cell.getNumericCellValue();
			break;

		case STRING:
			value = StringUtils.equalsIgnoreCase(cell.getStringCellValue().trim(), "null") ? null
					: cell.getStringCellValue().trim();
			break;

		case BLANK:
			break;

		case ERROR:
			break;

		case FORMULA:
			break;

		case _NONE:
			break;
		}

		return value;
	}
	
	/**
	 * This method will return the value from the provided cell as String
	 * 
	 * @param cell
	 * @return
	 */
	private String getValueFromCellasString(Cell cell) {
		Object cellValue = getValueFromCell(cell);
		return null == cellValue ? null : String.valueOf(getValueFromCell(cell));
	}
	
	/**
	 * This method is sed to handle common objects which are represented as [value1, value2]
	 * 
	 * @param coll
	 * @param clazz
	 * @param fieldValue
	 */
	@SuppressWarnings({ "unchecked", "rawtypes" })
	private void handleCommonObjects(Collection coll, Class<?> clazz, String fieldValue) {
		
		String[] splitValues = fieldValue.split(",");
		for (String value : splitValues) {
			value = value.trim();
			if (value.startsWith("[")) {
				value = value.substring(1, value.length());
			} else if (value.endsWith("]")) {
				value = value.substring(0, value.length() - 1);
			}

			coll.add(ConvertUtils.convert(value, clazz));
		}
	}
	
	/**
	 * This method handles the array
	 * 
	 * @param fieldType
	 * @param fieldValue
	 * @return
	 */
	private Object handleArray(Class<?> fieldType, String fieldValue) {
		
		Object objArr = null;
		Class<?> componentType = fieldType.getComponentType();
		Collection<?> coll = new ArrayList<>();

		if (fieldValue.contains("Sheet:")) {
			String sheetName = fieldValue.split(":")[1];
			convertSheetRowObjsToBeanFields(sheetName, coll, componentType);
			objArr = HelperUtils.convertObjectArrToCustomArr(componentType, coll.toArray());
		} else {
			handleCommonObjects(coll, componentType, fieldValue);
			objArr = HelperUtils.convertToPrimitiveArr(componentType, coll.toArray());
		}
		
		return objArr;
	}
	
	/**
	 * This method handles the collections
	 * 
	 * @param field
	 * @param fieldType
	 * @param fieldValue
	 * @return
	 * @throws NoSuchMethodException
	 * @throws IllegalAccessException
	 * @throws InvocationTargetException
	 * @throws InstantiationException
	 */
	private Collection<?> handleCollection(Field field, Class<?> fieldType, String fieldValue) 
			throws NoSuchMethodException, IllegalAccessException, InvocationTargetException, InstantiationException {
		Class<?> genericType = (Class<?>) ((ParameterizedType) field.getGenericType()).getActualTypeArguments()[0];
		Collection<?> coll = (Collection<?>) ConstructorUtils.invokeConstructor(fieldType);

		if (fieldValue.contains("Sheet:")) {
			String sheetName = fieldValue.split(":")[1];
			convertSheetRowObjsToBeanFields(sheetName, coll, genericType);
		} else {
			handleCommonObjects(coll, genericType, fieldValue);
		}
		
		return coll;
	}
	
	/**
	 * This method is used to create a new instance of the provided class
	 * 
	 * @param fieldType
	 * @return
	 * @throws NoSuchMethodException
	 * @throws SecurityException
	 * @throws InstantiationException
	 * @throws IllegalAccessException
	 * @throws IllegalArgumentException
	 * @throws InvocationTargetException
	 */
	private Object createInstance(Class<?> fieldType) 
			throws NoSuchMethodException, SecurityException, InstantiationException, IllegalAccessException, 
			IllegalArgumentException, InvocationTargetException {
		Constructor<?> ctor = fieldType.getDeclaredConstructor();
		ctor.setAccessible(true); 
		return ctor.newInstance();
	}
	
}
