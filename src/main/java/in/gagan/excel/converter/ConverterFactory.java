package in.gagan.excel.converter;

import in.gagan.excel.util.HelperUtils;

/**
 * Factory class to provide the instance of requested excel converter
 * 
 * @author gaganthind
 *
 */
public class ConverterFactory {

	private ConverterFactory() {
		/** Private Constructor */
	}

	/**
	 * Factory instance
	 * 
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

}
