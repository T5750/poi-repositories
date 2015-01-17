package template;

import java.util.Date;
import java.util.HashMap;
import java.util.Map;

//参考地址：http://mylfd.iteye.com/blog/1982101
public class TestTemplate {
	public static void main(String[] args) {
		ExcelTemplate excel = ExcelTemplate.getInstance().readTemplatePath(
				"D:/template.xls");
		excel.creatNewRow();
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("bbb");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.creatNewRow();
		excel.createNewCol("ccc");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.creatNewRow();
		excel.createNewCol("ddd");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.creatNewRow();
		excel.createNewCol("eee");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("555");
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "拉斯维加斯");
		datas.put("date", new Date().toString());
		datas.put("department", "百合科技人事部");
		excel.replaceFind(datas);
		excel.insertSer();
		excel.writeToFile("D:/poi.xls");
	}
}
