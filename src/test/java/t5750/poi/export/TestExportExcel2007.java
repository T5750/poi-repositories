package t5750.poi.export;

import java.io.File;

import t5750.poi.util.ExcelExportUtil;
import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

public class TestExportExcel2007 {
	public static void main(String[] args) {
		String filePath = TestUtil.DOC_PATH + File.separator
				+ Globals.EXPORT_2007;
		ExcelExportUtil.export2007(filePath);
		System.out.println(Globals.EXPORT_2007 + TestUtil.SUCCESS);
	}
}
