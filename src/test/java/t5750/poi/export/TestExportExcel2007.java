package t5750.poi.export;

import t5750.poi.util.ExcelExportUtil;
import t5750.poi.util.Globals;

public class TestExportExcel2007 {
	public static void main(String[] args) {
		String filePath = "D:/" + Globals.EXPORT_2007;
		ExcelExportUtil.export2007(filePath);
	}
}
