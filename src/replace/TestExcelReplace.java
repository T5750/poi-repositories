package replace;

import java.util.ArrayList;
import java.util.List;

//参考地址：http://yaoh6688.iteye.com/blog/1152273
public class TestExcelReplace {
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");

	public static void main(String[] args) {
		List<ExcelReplaceDataVO> datas = new ArrayList<ExcelReplaceDataVO>();
		// 找到第14行第2列的company，用"XXX有限公司"替换掉company
		ExcelReplaceDataVO vo1 = new ExcelReplaceDataVO();
		vo1.setRow(13);
		vo1.setColumn(1);
		vo1.setKey("company");
		vo1.setValue("XXX有限公司");
		// 找到第5行第2列的content，用"aa替换的内容aa"替换掉content
		ExcelReplaceDataVO vo2 = new ExcelReplaceDataVO();
		vo2.setRow(4);
		vo2.setColumn(1);
		vo2.setKey("content");
		vo2.setValue("aa替换的内容aa");
		datas.add(vo1);
		datas.add(vo2);
		// d:\\template.xls为Excel模板文件，d:\\test.xls为程序根据Excel模板文件生成的新文件
		ExcelUtil.replaceModel(datas, "d:\\replaceTemplate.xls", "d:\\eplace.xls");
		// String folderPath = "/home" + FILE_SEPARATOR;
		// ExcelUtil.replaceModel(datas, folderPath + "template.xls", folderPath
		// + "test.xls");
	}
}
