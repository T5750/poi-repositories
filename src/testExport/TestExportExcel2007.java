package testExport;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestExportExcel2007 {
	private static void createEXCEL() {
		try {
			// 输出流
			OutputStream os = new FileOutputStream("D:/export2007_"
					+ System.currentTimeMillis() + ".xlsx");
			// 工作区
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("test");
			for (int i = 0; i < 1000; i++) {
				// 创建第一个sheet
				// 生成第一行
				XSSFRow row = sheet.createRow(i);
				// 给这一行的第一列赋值
				row.createCell(0).setCellValue("column1");
				// 给这一行的第一列赋值
				row.createCell(1).setCellValue("column2");
				System.out.println(i);
			}
			// 写文件
			wb.write(os);
			// 关闭输出流
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		TestExportExcel2007.createEXCEL();
	}
}
