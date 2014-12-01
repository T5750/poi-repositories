package replace;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class ExcelUtil {
	/**
	 * 替换Excel模板文件内容
	 * 
	 * @param datas
	 *            文档数据
	 * @param sourceFilePath
	 *            Excel模板文件路径
	 * @param targetFilePath
	 *            Excel生成文件路径
	 */
	public static boolean replaceModel(List<ExcelReplaceDataVO> datas,
			String sourceFilePath, String targetFilePath) {
		boolean bool = true;
		try {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(
					sourceFilePath));
			HSSFWorkbook wb = new HSSFWorkbook(fs);
			HSSFSheet sheet = wb.getSheetAt(0);
			for (ExcelReplaceDataVO data : datas) {
				// 获取单元格内容
				HSSFRow row = sheet.getRow(data.getRow());
				HSSFCell cell = row.getCell((short) data.getColumn());
				String str = cell.getStringCellValue();
				// 替换单元格内容
				str = str.replace(data.getKey(), data.getValue());
				// 写入单元格内容
				cell.setCellType(HSSFCell.CELL_TYPE_STRING);
				// cell.setEncoding(HSSFCell.ENCODING_UTF_16);
				cell.setCellValue(str);
			}
			// 输出文件
			FileOutputStream fileOut = new FileOutputStream(targetFilePath);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			bool = false;
			e.printStackTrace();
		}
		return bool;
	}
}
