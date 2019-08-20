package t5750.poi.export;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

public class TestExcelStylingDemo {
	public static final String EXCEL_NAME = "testExcelStylingDemo";

	public static void main(String[] args) throws Exception {
		Workbook wb = new XSSFWorkbook();
		basedOnValue(wb.createSheet("Value Based formatting"));
		formatDuplicates(wb.createSheet("Duplicates formatting"));
		shadeAlt(wb.createSheet("Alternate rows"));
		expiryInNext30Days(wb.createSheet("Soon Expired Payments"));
		// Write the output to a file
		FileOutputStream out = new FileOutputStream(TestUtil.DOC_PATH
				+ File.separator + EXCEL_NAME + Globals.SUFFIX_XLSX);
		wb.write(out);
		out.close();
		System.out.println(EXCEL_NAME + Globals.SUFFIX_XLSX + TestUtil.SUCCESS);
	}

	// Cell value is in between a certain range
	static void basedOnValue(Sheet sheet) {
		// Creating some random values
		sheet.createRow(0).createCell(0).setCellValue(84);
		sheet.createRow(1).createCell(0).setCellValue(74);
		sheet.createRow(2).createCell(0).setCellValue(50);
		sheet.createRow(3).createCell(0).setCellValue(51);
		sheet.createRow(4).createCell(0).setCellValue(49);
		sheet.createRow(5).createCell(0).setCellValue(41);
		SheetConditionalFormatting sheetCF = sheet
				.getSheetConditionalFormatting();
		// Condition 1: Cell Value Is greater than 70 (Blue Fill)
		ConditionalFormattingRule rule1 = sheetCF
				.createConditionalFormattingRule(ComparisonOperator.GT, "70");
		PatternFormatting fill1 = rule1.createPatternFormatting();
		fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
		fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
		// Condition 2: Cell Value Is less than 50 (Green Fill)
		ConditionalFormattingRule rule2 = sheetCF
				.createConditionalFormattingRule(ComparisonOperator.LT, "50");
		PatternFormatting fill2 = rule2.createPatternFormatting();
		fill2.setFillBackgroundColor(IndexedColors.GREEN.index);
		fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
		CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:A6") };
		sheetCF.addConditionalFormatting(regions, rule1, rule2);
	}

	/**
	 * Use Excel conditional formatting to highlight duplicate entries in a
	 * column.
	 */
	static void formatDuplicates(Sheet sheet) {
		sheet.createRow(0).createCell(0).setCellValue("Code");
		sheet.createRow(1).createCell(0).setCellValue(4);
		sheet.createRow(2).createCell(0).setCellValue(3);
		sheet.createRow(3).createCell(0).setCellValue(6);
		sheet.createRow(4).createCell(0).setCellValue(3);
		sheet.createRow(5).createCell(0).setCellValue(5);
		sheet.createRow(6).createCell(0).setCellValue(8);
		sheet.createRow(7).createCell(0).setCellValue(0);
		sheet.createRow(8).createCell(0).setCellValue(2);
		sheet.createRow(9).createCell(0).setCellValue(8);
		sheet.createRow(10).createCell(0).setCellValue(6);
		SheetConditionalFormatting sheetCF = sheet
				.getSheetConditionalFormatting();
		// Condition 1: Formula Is =A2=A1 (White Font)
		ConditionalFormattingRule rule1 = sheetCF
				.createConditionalFormattingRule("COUNTIF($A$2:$A$11,A2)>1");
		FontFormatting font = rule1.createFontFormatting();
		font.setFontStyle(false, true);
		font.setFontColorIndex(IndexedColors.BLUE.index);
		CellRangeAddress[] regions = { CellRangeAddress.valueOf("A2:A11") };
		sheetCF.addConditionalFormatting(regions, rule1);
	}

	/**
	 * Use Excel conditional formatting to shade alternating rows on the
	 * worksheet
	 */
	static void shadeAlt(Sheet sheet) {
		SheetConditionalFormatting sheetCF = sheet
				.getSheetConditionalFormatting();
		// Condition 1: Formula Is =A2=A1 (White Font)
		ConditionalFormattingRule rule1 = sheetCF
				.createConditionalFormattingRule("MOD(ROW(),2)");
		PatternFormatting fill1 = rule1.createPatternFormatting();
		fill1.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.index);
		fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
		CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:Z100") };
		sheetCF.addConditionalFormatting(regions, rule1);
	}

	/**
	 * Highlight payments that are due in the next thirty days
	 */
	static void expiryInNext30Days(Sheet sheet) {
		CellStyle style = sheet.getWorkbook().createCellStyle();
		style.setDataFormat((short) BuiltinFormats.getBuiltinFormat("d-mmm"));
		sheet.createRow(0).createCell(0).setCellValue("Date");
		sheet.createRow(1).createCell(0).setCellFormula("TODAY()+29");
		sheet.createRow(2).createCell(0).setCellFormula("A2+1");
		sheet.createRow(3).createCell(0).setCellFormula("A3+1");
		for (int rownum = 1; rownum <= 3; rownum++)
			sheet.getRow(rownum).getCell(0).setCellStyle(style);
		SheetConditionalFormatting sheetCF = sheet
				.getSheetConditionalFormatting();
		// Condition 1: Formula Is =A2=A1 (White Font)
		ConditionalFormattingRule rule1 = sheetCF
				.createConditionalFormattingRule("AND(A2-TODAY()>=0,A2-TODAY()<=30)");
		FontFormatting font = rule1.createFontFormatting();
		font.setFontStyle(false, true);
		font.setFontColorIndex(IndexedColors.BLUE.index);
		CellRangeAddress[] regions = { CellRangeAddress.valueOf("A2:A4") };
		sheetCF.addConditionalFormatting(regions, rule1);
		sheet.getRow(0).createCell(1)
				.setCellValue("Dates within the next 30 days are highlighted");
	}
}
