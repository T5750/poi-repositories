package t5750.poi.util;

import java.io.File;

import javax.servlet.http.HttpServletRequest;

public class Globals {
	public static final String DOC = "doc";
	public static final String IMG = "img";
	public static final String DOC_PATH = ClassLoaderUtil.getClassPath()
			+ File.separator + DOC;
	public static final String IMG_PATH = ClassLoaderUtil.getClassPath()
			+ File.separator + IMG;
	public static final String SUFFIX_XLS = ".xls";
	public static final String SUFFIX_XLSX = ".xlsx";
	public static final String SUFFIX_PNG = ".png";
	public static final String EXPORT_STUDENT = "student" + SUFFIX_XLS;
	public static final String EXPORT_BOOK = "book" + SUFFIX_XLS;
	public static final String EXPORT_PRODUCT = "product" + SUFFIX_XLS;

	public static String getBasePath(HttpServletRequest request) {
		String path = request.getContextPath();
		String basePath = request.getScheme() + "://" + request.getServerName()
				+ ":" + request.getServerPort() + path + "/";
		return basePath;
	}
}