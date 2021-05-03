package t5750.poi.replace;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

import t5750.poi.command.ExcelReplaceDataVO;
import t5750.poi.util.ExcelReplaceUtil;
import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

/**
 * 参考地址：http://yaoh6688.iteye.com/blog/1152273
 */
public class TestExcelReplace {
	public static final String EXCEL_NAME = "replaceTemplate";

	public static void main(String[] args) throws FileNotFoundException {
		List<ExcelReplaceDataVO> datas = new ArrayList<ExcelReplaceDataVO>();
		// 找到第14行第2列的company，用"XXX有限公司"替换掉company
		ExcelReplaceDataVO voCompany = new ExcelReplaceDataVO();
		voCompany.setRow(13);
		voCompany.setColumn(1);
		voCompany.setKey("company");
		voCompany.setValue("XXX有限公司");
		// 找到第5行第2列的content，用"替换的内容"替换掉content
		ExcelReplaceDataVO voContent = new ExcelReplaceDataVO();
		voContent.setRow(4);
		voContent.setColumn(1);
		voContent.setKey("content");
		voContent.setValue("替换的内容");
		datas.add(voCompany);
		datas.add(voContent);
		// replaceTemplate.xls为Excel模板文件，d:\\replaceTemplate*.xls为程序根据Excel模板文件生成的新文件
		ExcelReplaceUtil.replaceModel(datas,
				TestUtil.DOC_PATH + File.separator + EXCEL_NAME
						+ Globals.SUFFIX_XLS,
				TestUtil.DOC_PATH + File.separator + EXCEL_NAME
						+ System.currentTimeMillis() + Globals.SUFFIX_XLS);
		System.out.println(EXCEL_NAME + Globals.SUFFIX_XLSX + TestUtil.SUCCESS);
	}
}
