package t5750.poi.template;

import java.io.File;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import t5750.poi.util.DateUtil;
import t5750.poi.util.ExcelTemplateUtil;
import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

/**
 * 参考地址：http://mylfd.iteye.com/blog/1982101
 */
public class TestTemplate {
	public static final String EXCEL_NAME = "template";

	public static void main(String[] args) {
		ExcelTemplateUtil excel = ExcelTemplateUtil.getInstance()
				.readTemplatePath(
						TestUtil.DOC_PATH + File.separator + EXCEL_NAME
								+ Globals.SUFFIX_XLS);
		for (int i = 0; i < 5; i++) {
			excel.creatNewRow();
			excel.createNewCol("Col" + i);
			excel.createNewCol(i);
			excel.createNewCol(i);
		}
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "Apache POI");
		datas.put("content", "the Java API for Microsoft Documents");
		datas.put("date", DateUtil.format(new Date()));
		excel.replaceFind(datas);
		excel.insertSer();
		excel.writeToFile(TestUtil.DOC_PATH + File.separator + EXCEL_NAME
				+ System.currentTimeMillis() + Globals.SUFFIX_XLS);
		System.out.println(EXCEL_NAME + Globals.SUFFIX_XLSX + TestUtil.SUCCESS);
	}
}
