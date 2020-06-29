package in.gagan.excel.util;

import java.lang.reflect.Array;
import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.Arrays;

import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.ClassUtils;

/**
 * This class contains helper functions
 * 
 * @author gaganthind
 *
 */
public abstract class HelperUtils {
	
	private HelperUtils() { /** Private Constructor */ }
	
	/**
	 * Check whether the class exists and is not initialized to null
	 * 
	 * @param obj
	 * @throws ClassNotFoundException
	 */
	public static void checkIfClassExists(Object obj) throws ClassNotFoundException {
		
		// Pre-Work: Check for the null initialization and class exists
		if (null == obj) {
			throw new NullPointerException("EDCaseObjectToExcelUtil: Object parameter is null");
		}

		try {
			// Check if the class to be converted is present in runtime
			Class.forName(obj.getClass().getName());
		} catch (ClassNotFoundException e) {
			throw new ClassNotFoundException("No class found for " + obj.getClass().getName()
					+ " in runtime: " + e.getMessage());
		}
	}

	/**
	 * Check if the provided class is a complex class
	 * 
	 * @param genericType
	 * @return
	 */
	public static boolean isComplexObject(Class<?> genericType) {
		return !genericType.isEnum() && !ClassUtils.isPrimitiveOrWrapper(genericType) && !genericType.isAssignableFrom(String.class);
	}

	/**
	 * Check if provided field is static final
	 *
	 * @param field
	 * @return
	 */
	public static boolean isStaticFinalField(Field field) {
		return Modifier.isStatic(field.getModifiers()) && Modifier.isFinal(field.getModifiers());
	}

	/**
	 * Convert object to primitive arr
	 *
	 * @param type
	 * @param inputObj
	 * @return
	 */
	public static Object convertToPrimitiveArr(Class<?> type, Object inputObj) {
		return ConvertUtils.convert(inputObj, Array.newInstance(type, 0).getClass());
	}

	/**
	 * Convert Object arr to custom class arr
	 *
	 * @param type
	 * @param inputObj
	 * @return
	 */
	public static Object convertObjectArrToCustomArr(Class<?> type, Object[] inputObj) {
		Object[] objArr = (Object[]) Array.newInstance(type, inputObj.length);

		for (int i = 0; i < inputObj.length; i++) {
			objArr[i] = type.cast(inputObj[i]);
		}

		return objArr;
	}

	/**
	 * Write array object as String
	 *
	 * @param obj
	 * @return
	 */
	public static String writeArrayObjectAsString(Object obj) {
		if (obj instanceof Object[]) {
			return Arrays.toString((Object[]) obj);
		} else {
			int length = Array.getLength(obj);
			Object[] objArr = new Object[length];
			for (int i = 0; i < length; i++) {
				objArr[i] = Array.get(obj, i);
			}
			
			return Arrays.toString(objArr);
		}
	}
}
