package t5750.poi.export;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

public class TestExportExcel2007 {
	private static final Pattern utfPtrn = Pattern
			.compile("_(x[0-9A-Fa-f]{4}_)");
	private static final String UNICODE_CHARACTER_LOW_LINE = "_x005F_";

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
			cell.setCellValue(TestUtil.RICH_TEXT_STRINGS[0]
					+ escape(TestUtil.REGEX + TestUtil.RICH_TEXT_STRINGS[1]
							+ TestUtil.REGEX) + TestUtil.RICH_TEXT_STRINGS[2]);
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

	/**
	 * https://bz.apache.org/bugzilla/show_bug.cgi?id=57008
	 */
	public static String escape(final String value) {
		// https://stackoverflow.com/questions/48222502/xssfcell-in-apache-poi-encodes-certain-character-sequences-as-unicode-character
		if (value == null)
			return null;
		StringBuffer buf = new StringBuffer();
		Matcher m = utfPtrn.matcher(value);
		int idx = 0;
		while (m.find()) {
			int pos = m.start();
			if (pos > idx) {
				buf.append(value.substring(idx, pos));
			}
			buf.append(UNICODE_CHARACTER_LOW_LINE + m.group(1));
			idx = m.end();
		}
		buf.append(value.substring(idx));
		return buf.toString();
	}
}