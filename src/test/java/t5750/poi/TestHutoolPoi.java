package t5750.poi;

import java.util.List;
import java.util.Map;

import org.junit.Test;

import cn.hutool.core.collection.CollUtil;
import cn.hutool.core.date.DateUtil;
import cn.hutool.core.lang.Console;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelReader;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;

public class TestHutoolPoi {
	public static final String FILE_PATH = "D:/TestHutoolPoi.xlsx";

	/**
	 * <a href=
	 * "https://doc.hutool.cn/pages/BigExcelWriter/">Excel大数据生成-BigExcelWriter</a>
	 */
	@Test
	public void testBigExcelWriter() {
		List<?> row1 = CollUtil.newArrayList("aa", "bb", "cc", "dd",
				DateUtil.date(), 3.22676575765);
		List<?> row2 = CollUtil.newArrayList("aa1", "bb1", "cc1", "dd1",
				DateUtil.date(), 250.7676);
		List<?> row3 = CollUtil.newArrayList("aa2", "bb2", "cc2", "dd2",
				DateUtil.date(), 0.111);
		List<?> row4 = CollUtil.newArrayList("aa3", "bb3", "cc3", "dd3",
				DateUtil.date(), 35);
		List<?> row5 = CollUtil.newArrayList("aa4", "bb4", "cc4", "dd4",
				DateUtil.date(), 28.00);
		List<List<?>> rows = CollUtil.newArrayList(row1, row2, row3, row4,
				row5);
		BigExcelWriter writer = ExcelUtil.getBigWriter(FILE_PATH);
		// 一次性写出内容，使用默认样式
		writer.write(rows);
		// 关闭writer，释放内存
		writer.close();
	}

	/**
	 * <a href=
	 * "https://doc.hutool.cn/pages/ExcelReader/">Excel读取-ExcelReader</a>
	 */
	@Test
	public void testExcelReader() {
		ExcelReader reader = ExcelUtil.getReader(FILE_PATH);
		// 1.读取Excel中所有行和列，用二维列表表示
		List<List<Object>> readAllList = reader.read();
		System.out.println(readAllList);
		List<Map<String, Object>> readAllMap = reader.readAll();
		// 2.读取为Map列表，默认第一行为标题行，数据从第二行开始，一个Map表示一行，Map中的key为标题，value为标题对应的单元格值
		System.out.println(readAllMap);
		// 3.读取为Bean列表，Bean中的字段名为标题，字段值为标题对应的单元格值
		// List<Person> all = reader.readAll(Person.class);
	}

	/**
	 * <a href=
	 * "https://doc.hutool.cn/pages/Excel07SaxReader/">流方式读取Excel2007-Excel07SaxReader</a>
	 */
	@Test
	public void testExcel07SaxReader() {
		ExcelUtil.readBySax(FILE_PATH, "sheet1", createRowHandler());
		// reader方法的第二个参数是sheet的序号，-1表示读取所有sheet，0表示第一个sheet
		Excel07SaxReader reader = new Excel07SaxReader(createRowHandler());
		Excel07SaxReader excel07SaxReader = reader.read(FILE_PATH, 0);
	}

	private RowHandler createRowHandler() {
		return new RowHandler() {
			@Override
			public void handle(int sheetIndex, long rowIndex,
					List<Object> rowlist) {
				Console.log("[{}] [{}] {}", sheetIndex, rowIndex, rowlist);
			}
		};
	}
}
