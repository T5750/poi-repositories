package t5750.poi.export;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.test.context.junit4.SpringRunner;

import t5750.poi.PoiApplication;
import t5750.poi.util.ExcelExportUtil;
import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = PoiApplication.class, webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
public class TestExportExcel {
	@Test
	public void exportExcel() throws IOException {
		String fileName = "default";
		String docsPath;
		Resource resource = new ClassPathResource(Globals.DOC + File.separator
				+ fileName + Globals.SUFFIX_XLS);
		if (resource.exists()) {
			docsPath = resource.getFile().getPath();
		} else {
			String imagesPath = TestUtil.IMG_PATH + File.separator + "tomcat"
					+ Globals.SUFFIX_PNG;
			ExcelExportUtil.export2003(imagesPath, TestUtil.DOC_PATH);
			docsPath = TestUtil.DOC_PATH + File.separator + Globals.EXPORT_BOOK;
		}
		System.out.println(docsPath);
	}

	@Test
	public void exportExcelWithStyle() {
		try {
			String filePath = TestUtil.DOC_PATH + File.separator
					+ Globals.EXPORT_PRODUCT;
			OutputStream os = new FileOutputStream(filePath);
			HSSFWorkbook wb = new HSSFWorkbook();
			HSSFSheet sheet = wb.createSheet(Globals.SHEETNAME);
			HSSFRichTextString richString = new HSSFRichTextString(
					TestUtil.RICH_TEXT_STRING);
			HSSFFont font = wb.createFont();
			font.setColor(IndexedColors.BLUE.index);
			richString.applyFont(font);
			sheet.createRow(0).createCell(0).setCellValue(richString);
			wb.write(os);
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}