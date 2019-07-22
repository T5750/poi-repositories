package t5750.poi.replace;

import java.util.ArrayList;
import java.util.List;

import t5750.poi.command.ExcelReplaceDataVO;
import t5750.poi.util.ExcelReplaceUtil;
import t5750.poi.util.Globals;

/**
 * 参考地址：http://yaoh6688.iteye.com/blog/1152273
 */
public class TestExcelReplace {
	public static void main(String[] args) {
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
		// d:\\template.xls为Excel模板文件，d:\\test.xls为程序根据Excel模板文件生成的新文件
		ExcelReplaceUtil.replaceModel(datas, "d:\\replaceTemplate"
				+ Globals.SUFFIX_XLS,
				"d:\\replaceTemplate" + System.currentTimeMillis()
						+ Globals.SUFFIX_XLS);
		// String folderPath = "/home" + FILE_SEPARATOR;
		// ExcelUtil.replaceModel(datas, folderPath + "template.xls", folderPath
		// + "test.xls");
	}
}
