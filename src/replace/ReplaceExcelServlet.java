package replace;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class ReplaceExcelServlet extends HttpServlet {
	public static final String FILE_SEPARATOR = System.getProperties()
			.getProperty("file.separator");

	@Override
	protected void doPost(HttpServletRequest request,
			HttpServletResponse response) throws ServletException, IOException {
		doGet(request, response);
	}

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String docsPath = request.getSession().getServletContext()
				.getRealPath("docs");
		String templateName = "replaceTemplate.xls";// Excel模版
		String templatePath = docsPath + FILE_SEPARATOR + templateName;
		String fileName = "replace.xls";// 导出Excel文件名
		String filePath = docsPath + FILE_SEPARATOR + fileName;
		List<ExcelReplaceDataVO> datas = new ArrayList<ExcelReplaceDataVO>();
		// 找到第14行第2列的company，用"XXX有限公司"替换掉company
		ExcelReplaceDataVO vo1 = new ExcelReplaceDataVO();
		vo1.setRow(13);
		vo1.setColumn(1);
		vo1.setKey("company");
		vo1.setValue("XXX有限公司");
		// 找到第5行第2列的content，用"aa替换的内容aa"替换掉content
		ExcelReplaceDataVO vo2 = new ExcelReplaceDataVO();
		vo2.setRow(4);
		vo2.setColumn(1);
		vo2.setKey("content");
		vo2.setValue("aa替换的内容aa");
		datas.add(vo1);
		datas.add(vo2);
		ExcelUtil.replaceModel(datas, templatePath, filePath);
		download(filePath, response);
	}

	private void download(String path, HttpServletResponse response) {
		try {
			// path是指欲下载的文件的路径。
			File file = new File(path);
			// 取得文件名。
			String filename = file.getName();
			// 以流的形式下载文件。
			InputStream fis = new BufferedInputStream(new FileInputStream(path));
			byte[] buffer = new byte[fis.available()];
			fis.read(buffer);
			fis.close();
			// 清空response
			response.reset();
			// 设置response的Header
			response.addHeader("Content-Disposition", "attachment;filename="
					+ new String(filename.getBytes()));
			response.addHeader("Content-Length", "" + file.length());
			OutputStream toClient = new BufferedOutputStream(
					response.getOutputStream());
			response.setContentType("application/vnd.ms-excel;charset=gb2312");
			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (IOException ex) {
			ex.printStackTrace();
		}
	}
}