package t5750.poi.export;

import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import t5750.poi.util.Globals;

public class TestExportExcel2007 {
	private static void createExcel() {
		try {
			// 输出流
			OutputStream os = new FileOutputStream("D:/export2007_"
					+ System.currentTimeMillis() + Globals.SUFFIX_XLSX);
			// 工作区
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet("test");
			for (int i = 0; i < 1000; i++) {
				// 创建第一个sheet
				// 生成第一行
				XSSFRow row = sheet.createRow(i);
				// 给这一行的第一列赋值
				row.createCell(0).setCellValue("column" + i);
				// 给这一行的第一列赋值
				row.createCell(1).setCellValue("column" + i);
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
		TestExportExcel2007.createExcel();
	}
}
