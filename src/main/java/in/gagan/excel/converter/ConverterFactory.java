package in.gagan.excel.converter;

import in.gagan.excel.util.HelperUtils;

/**
 * Factory class to provide the instance of requested excel converter
 * 
 * @author gaganthind
 *
 */
public class ConverterFactory {

	private ConverterFactory() { }

	/**
	 * Factory instance
	 * 
	 * @params converterType the converter type to use
	 * @params obj Object to be used for converting to excel or for writing back
	 * @params path Path where the excel file should be saved/retreived
	 * @return Converter Object
	 * @throws ClassNotFoundException
	 */
	public static Converter getConverter(ConverterType converterType, Object obj, final String path)
			throws ClassNotFoundException {

		// Pre-Work
		HelperUtils.checkIfClassExists(obj);

		if (ConverterType.POI_CONVERTER == converterType) {
			return new PoiConverter(obj, path);
		}

		return null;

	}
	
	/**
	 * Factory instance
	 * 
	 * @params obj Object to be used for converting to excel or for writing back
	 * @params path Path where the excel file should be saved/retreived
	 * @return Converter Object
	 * @throws ClassNotFoundException
	 */
	public static Converter getConverter(Object obj, final String path) throws ClassNotFoundException {
		return getConverter(ConverterType.POI_CONVERTER, obj, path)
	}

}
