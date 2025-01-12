package t5750.poi.util;

import java.io.IOException;
import java.net.URLEncoder;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import cn.hutool.core.date.DateUtil;
import cn.hutool.core.io.IoUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import t5750.poi.domain.SheetDTO;

/**
 * <a href=
 * "https://blog.csdn.net/ZLK1142/article/details/106531246">Hutool导出Excel，导多个Sheet页</a>
 */
public class HuExcelUtil {
	/**
	 * 导出多个 Sheet 页
	 * 
	 * @param response
	 * @param sheetList
	 *            页数据
	 * @param fileName
	 *            文件名
	 */
	public static void exportExcel(HttpServletResponse response,
			List<SheetDTO> sheetList, String fileName) {
		ExcelWriter bigWriter = ExcelUtil.getBigWriter();
		// 重命名第一个Sheet的名称，不然会默认多出一个Sheet1的页
		bigWriter.renameSheet(0, sheetList.get(0).getSheetName());
		for (SheetDTO sheet : sheetList) {
			// 指定要写出的 Sheet 页
			bigWriter.setSheet(sheet.getSheetName());
			List<Integer> columnWidth = sheet.getColumnWidth();
			if (columnWidth == null
					|| columnWidth.size() != sheet.getFieldAndAlias().size()) {
				// 设置默认宽度
				for (int i = 0; i < sheet.getFieldAndAlias().size(); i++) {
					bigWriter.setColumnWidth(i, 25);
				}
			} else {
				// 设置自定义宽度
				for (int i = 0; i < columnWidth.size(); i++) {
					bigWriter.setColumnWidth(i, columnWidth.get(i));
				}
			}
			// 设置字段和别名
			bigWriter.setHeaderAlias(sheet.getFieldAndAlias());
			// 设置只导出有别名的字段
			bigWriter.setOnlyAlias(true);
			// 设置默认行高
			bigWriter.setDefaultRowHeight(18);
			// 设置冻结行
			bigWriter.setFreezePane(1);
			// 一次性写出内容，使用默认样式，强制输出标题
			bigWriter.write(sheet.getCollection(), true);
			// 设置所有列为自动宽度，不考虑合并单元格
			// bigWriter.autoSizeColumnAll();
		}
		ServletOutputStream out = null;
		try {
			// response为HttpServletResponse对象
			response.setContentType("application/vnd.ms-excel;charset=utf-8");
			response.setHeader("Content-Disposition",
					"attachment;filename=" + URLEncoder.encode(
							fileName + DateUtil.today() + ".xlsx", "UTF-8"));
			out = response.getOutputStream();
			bigWriter.flush(out, true);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			// 关闭writer，释放内存
			bigWriter.close();
		}
		// 此处记得关闭输出Servlet流
		IoUtil.close(out);
	}
}