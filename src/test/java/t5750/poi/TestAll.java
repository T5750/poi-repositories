package t5750.poi;

import org.apache.poi.ss.examples.CalendarDemo;

import t5750.poi.export.TestExcelFormulaDemo;
import t5750.poi.export.TestExcelStylingDemo;
import t5750.poi.export.TestExportExcel2007;
import t5750.poi.export.TestWriteExcelDemo;
import t5750.poi.read.TestReadExcelDemo;
import t5750.poi.replace.TestExcelReplace;
import t5750.poi.template.TestTemplate;

public class TestAll {
	public static void main(String[] args) throws Exception {
		TestExportExcel2007.main(args);
		System.out.println();
		TestTemplate.main(args);
		System.out.println();
		TestExcelReplace.main(args);
		System.out.println();
		TestWriteExcelDemo.main(args);
		System.out.println();
		TestReadExcelDemo.main(args);
		System.out.println();
		TestExcelFormulaDemo.main(args);
		System.out.println();
		TestExcelStylingDemo.main(args);
		System.out.println();
		CalendarDemo.main(args);
	}
}
