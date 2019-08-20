package t5750.poi.export;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

public class TestExportExcel2007 {
	public static void main(String[] args) {
		String filePath = TestUtil.DOC_PATH + File.separator
				+ Globals.EXPORT_2007;
		// ExcelExportUtil.export2007(filePath);
		export2007WithStyle(filePath);
		System.out.println(Globals.EXPORT_2007 + TestUtil.SUCCESS);
	}

	/**
	 * XSSFRichTextString.utfDecode()<br/>
	 * value.contains("_x")<br/>
	 * Pattern.compile("_x([0-9A-Fa-f]{4})_");
	 */
	private static void export2007WithStyle(String filePath) {
		try {
			OutputStream os = new FileOutputStream(filePath);
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet(Globals.SHEETNAME);
			XSSFCell cell = sheet.createRow(0).createCell(0);
			cell.setCellValue(TestUtil.RICH_TEXT_STRING);
			CellStyle style = sheet.getWorkbook().createCellStyle();
			XSSFFont font = wb.createFont();
			font.setColor(IndexedColors.BLUE.index);
			style.setFont(font);
			cell.setCellStyle(style);
			// richString.applyFont(font);
			wb.write(os);
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}