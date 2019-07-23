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
	public static final String EXPORT_STUDENT = "student"
			+ System.currentTimeMillis() + SUFFIX_XLS;
	public static final String EXPORT_BOOK = "book"
			+ System.currentTimeMillis() + SUFFIX_XLS;
	public static final String EXPORT_2007 = "export2007中文_"
			+ System.currentTimeMillis() + SUFFIX_XLSX;
	public static final String EXPORT_PRODUCT = "product"
			+ System.currentTimeMillis() + SUFFIX_XLS;
	public static final String SHEETNAME = "测试POI导出EXCEL文档";

	public static String getBasePath(HttpServletRequest request) {
		String path = request.getContextPath();
		String basePath = request.getScheme() + "://" + request.getServerName()
				+ ":" + request.getServerPort() + path + "/";
		return basePath;
	}
}