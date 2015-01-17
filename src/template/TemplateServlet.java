package template;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

public class TemplateServlet extends HttpServlet {
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
		String templateName = "template.xls";// Excel模版
		String templatePath = docsPath + FILE_SEPARATOR + templateName;
		String fileName = "testTemplate.xls";// 导出Excel文件名
		String filePath = docsPath + FILE_SEPARATOR + fileName;
		ExcelTemplate excel = ExcelTemplate.getInstance().readTemplatePath(
				templatePath);
		excel.creatNewRow();
		excel.createNewCol("aaa");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.createNewCol("111");
		excel.creatNewRow();
		excel.createNewCol("bbb");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.createNewCol("222");
		excel.creatNewRow();
		excel.createNewCol("ccc");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.createNewCol("333");
		excel.creatNewRow();
		excel.createNewCol("ddd");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.createNewCol("444");
		excel.creatNewRow();
		excel.createNewCol("eee");
		excel.createNewCol("555");
		excel.createNewCol("555");
		excel.createNewCol("555");
		Map<String, String> datas = new HashMap<String, String>();
		datas.put("title", "拉斯维加斯");
		datas.put("date", new Date().toString());
		datas.put("department", "百合科技人事部");
		excel.replaceFind(datas);
		excel.insertSer();
		excel.writeToFile(filePath);
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