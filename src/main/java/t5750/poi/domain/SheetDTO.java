package t5750.poi.domain;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;
import java.util.Map;

/**
 * Excel - Sheet页
 */
public class SheetDTO implements Serializable {
	private static final long serialVersionUID = 1L;
	/**
	 * sheet页名称
	 */
	private String sheetName;
	/**
	 * 字段和别名，如果使用这个，properties 和 titles可以不用处理 Map<字段, 别名> 如：Map<"name", "姓名">
	 */
	private Map<String, String> fieldAndAlias;
	/**
	 * 列宽<br/>
	 * 设置列宽时必须每个字段都设置才生效（columnWidth.size = fieldAndAlias.size）
	 */
	private List<Integer> columnWidth;
	/**
	 * 数据集
	 */
	private Collection<?> collection;

	public SheetDTO() {
	}

	/**
	 * @param sheetName
	 *            sheet页名称
	 * @param fieldAndAlias
	 *            字段和别名
	 * @param collection
	 *            数据集
	 */
	public SheetDTO(String sheetName, Map<String, String> fieldAndAlias,
			Collection<?> collection) {
		super();
		this.sheetName = sheetName;
		this.fieldAndAlias = fieldAndAlias;
		this.collection = collection;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public Map<String, String> getFieldAndAlias() {
		return fieldAndAlias;
	}

	public void setFieldAndAlias(Map<String, String> fieldAndAlias) {
		this.fieldAndAlias = fieldAndAlias;
	}

	public List<Integer> getColumnWidth() {
		return this.columnWidth;
	}

	public void setColumnWidth(List<Integer> columnWidth) {
		this.columnWidth = columnWidth;
	}

	public Collection<?> getCollection() {
		return collection;
	}

	public void setCollection(Collection<?> collection) {
		this.collection = collection;
	}
}