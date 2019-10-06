package t5750.poi.util;

public class TestUtil {
	/**
	 * poi\target\classes\doc
	 */
	public static final String DOC_PATH = Globals.DOC_PATH.replace(
			"test-classes", "classes");
	public static final String IMG_PATH = Globals.IMG_PATH.replace(
			"test-classes", "classes");
	public static final String SUCCESS = " written successfully...";
	public static final String RICH_TEXT_STRING = "o9qAGtxS7K25ItIvnNXd_xBcCe_0";
	public static final String REGEX = "_";
	public static final String[] RICH_TEXT_STRINGS = RICH_TEXT_STRING
			.split(REGEX);
}