package t5750.poi.util;

import java.io.*;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

/**
 * 用于处理Excel写操作
 */
public class ExcelTemplateUtil {
	// 开始位置
	public final static String TEMPLATE_NAME = "datastart";
	// 样式,用于要标志自定义的样式的列.
	public final static String STYLE = "style";
	// 默认样式
	public final static String DEFALULT_STYLE = "defaultStyle";
	// 序号,确定Excel是否需要样式.
	public final static String SERNUMS = "sernums";
	private int initRowIndex; // 初始行
	private int initColIndex; // 初始列
	private int curRowIndex; // 当前行
	private int curColIndex; // 当前列
	private int lastRowInex; // 最后一行
	private float defaultHeight; // 默认行高
	private int serColIndex; // 序号行.
	private Workbook workbook = null;
	private Sheet sheet = null;
	private Row curRow = null; // 当前行
	// 样式
	private Map<Integer, CellStyle> styles = null;
	// 默认样式
	private CellStyle defaultStyle = null;
	// 使用单例
	private static ExcelTemplateUtil excel = new ExcelTemplateUtil();

	private ExcelTemplateUtil() {
	}

	public static ExcelTemplateUtil getInstance() {
		return excel;
	}

	/**
	 * 读取模板(从classpath中)
	 *
	 * @param path
	 *            模板路径
	 * @return ExcelTemplate
	 */
	public ExcelTemplateUtil readTemplateClassPath(String calsspath) {
		try {
			workbook = WorkbookFactory.create(
					ExcelTemplateUtil.class.getResourceAsStream(calsspath));
			initTemplate();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板格式有错!请检查.");
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板文件不存在!请检查.");
		}
		return this;
	}

	/**
	 * 读取模板(从指定路径)
	 *
	 * @param path
	 *            模板路径
	 * @return ExcelTemplate
	 */
	public ExcelTemplateUtil readTemplatePath(String path) {
		try {
			workbook = WorkbookFactory.create(new File(path));
			initTemplate();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板格式有错!请检查.");
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板文件不存在!请检查.");
		}
		return this;
	}

	public ExcelTemplateUtil readTemplatePath(InputStream is) {
		try {
			workbook = WorkbookFactory.create(is);
			initTemplate();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板格式有错!请检查.");
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("读取模板文件不存在!请检查.");
		}
		return this;
	}

	/**
	 * 创建新的一行
	 */
	public void creatNewRow() {
		// curRowIndex != initRowIndxe : 当前行本身是存在的,所以下移多余.
		if (lastRowInex > curRowIndex && curRowIndex != initRowIndex) {
			sheet.shiftRows(curRowIndex, lastRowInex, 1, true, true); // 有的模板最后可能是日期或者姓名之类的非数据.所以要移动行.
			lastRowInex++;
		}
		curRow = sheet.createRow(curRowIndex);
		curRow.setHeightInPoints(defaultHeight);
		curRowIndex++;
		curColIndex = initColIndex;
	}

	public void createNewCol(String value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	public void createNewCol(double value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	public void createNewCol(boolean value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	public void createNewCol(Date value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	public void createNewCol(Calendar value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	public void createNewCol(RichTextString value) {
		Cell cell = curRow.createCell(curColIndex);
		setStyle(cell);
		cell.setCellValue(value);
		curColIndex++;
	}

	/**
	 * 根据#xxx替换模板中的其它样式.
	 *
	 * @param datas
	 *            要替换的数据
	 */
	public void replaceFind(Map<String, String> datas) {
		if (datas == null)
			return;
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING)
					continue;
				String value = cell.getStringCellValue().trim();
				if (value.startsWith("#")) {
					if (datas.containsKey(value.substring(1))) {
						cell.setCellValue(datas.get(value.substring(1)));
					}
				}
			}
		}
	}

	/**
	 * 插入序号
	 */
	public void insertSer() {
		int index = 1;
		Row row = null;
		Cell cell = null;
		for (int i = initRowIndex; i < curRowIndex; i++) {
			row = sheet.getRow(i);
			cell = row.createCell(serColIndex);
			setStyle(cell);
			cell.setCellValue(index++);
		}
	}

	/**
	 * 输出文件,根据路径
	 *
	 * @param path
	 *            路径
	 */
	public void writeToFile(String path) {
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream(path);
			workbook.write(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			throw new RuntimeException("找不到文件!请检查.");
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("文件输出异常!请检查.");
		} finally {
			try {
				if (fos != null) {
					fos.close();
					fos = null;
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 输出文件,根据流输出
	 *
	 * @param stream
	 *            OutputStream
	 */
	public void writeToStream(OutputStream stream) {
		try {
			workbook.write(stream);
		} catch (IOException e) {
			e.printStackTrace();
			throw new RuntimeException("文件输出异常!请检查.");
		} finally {
			try {
				if (stream != null) {
					stream.close();
					stream = null;
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	/**
	 * 初始化模板
	 */
	private void initTemplate() {
		sheet = workbook.getSheetAt(0);
		styles = new HashMap<Integer, CellStyle>();
		initConfigData();
		lastRowInex = sheet.getLastRowNum();
	}

	/**
	 * defaultStyles:获得默认样式(如果默认样式没有则使用开始样式) styles:获取自定义样式
	 *
	 */
	private void initConfigData() {
		for (Row row : sheet) {
			for (Cell cell : row) {
				if (cell.getCellType() != Cell.CELL_TYPE_STRING)
					continue;
				String value = cell.getStringCellValue().trim();
				// 获取开始位置,初始化数据
				if (TEMPLATE_NAME.equals(value)) {
					initRowIndex = cell.getRowIndex();
					initColIndex = cell.getColumnIndex();
					curRowIndex = initRowIndex;
					curColIndex = initColIndex;
					defaultHeight = row.getHeightInPoints();
					if (defaultStyle == null)
						defaultStyle = cell.getCellStyle();
				}
				// 获取defaultStyles,无论如何,当有设置defaultStyles都设置为defaultStyles
				if (DEFALULT_STYLE.equals(value))
					defaultStyle = cell.getCellStyle();
				// 获取自定义样式的列
				if (STYLE.equals(value)) {
					styles.put(cell.getColumnIndex(), cell.getCellStyle());
				}
				// 获取样式所在的列
				if (SERNUMS.equals(value))
					serColIndex = cell.getColumnIndex();
			}
		}
	}

	/**
	 * 设置样式
	 *
	 * @param cell
	 *            Cell
	 */
	private void setStyle(Cell cell) {
		// 当前列存在自定义样式时使用自定义样式,否则使用默认样式.
		if (styles.containsKey(curColIndex)) {
			cell.setCellStyle(styles.get(curColIndex));
		} else {
			cell.setCellStyle(defaultStyle);
		}
	}
}
