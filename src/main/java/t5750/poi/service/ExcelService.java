package t5750.poi.service;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

public interface ExcelService {
	List<List<Object>> readExcel(File file) throws IOException;

	List<List<Object>> readExcel(InputStream is, String suffix)
			throws IOException;

	void download(String filename, InputStream is, HttpServletResponse response)
			throws IOException;

	void download(String filename, String path, HttpServletResponse response)
			throws IOException;

	String export2003(String fileName, HttpServletResponse response)
			throws IOException;

	String export2007(String fileName, HttpServletResponse response)
			throws IOException;

	String template(String fileName, HttpServletResponse response)
			throws IOException;

	String replace(String fileName, HttpServletResponse response)
			throws IOException;
}
