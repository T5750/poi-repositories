package t5750.poi.controller;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import t5750.poi.service.ExcelService;
import t5750.poi.util.Globals;

@RestController
@RequestMapping("/excel")
public class ExcelController {
	@Autowired
	private ExcelService excelService;
	@Autowired
	private HttpServletResponse response;

	@RequestMapping(value = "/read/{fileName}", method = RequestMethod.GET)
	public List read(@PathVariable String fileName) throws IOException {
		Resource resource = new ClassPathResource(
				Globals.DOC + File.separator + fileName + Globals.SUFFIX_XLS);
		InputStream is = resource.getInputStream();
		List<List<Object>> list = excelService.readExcel(is,
				Globals.SUFFIX_XLS);
		return list;
	}

	@RequestMapping(value = "/export/{fileName}", method = RequestMethod.GET)
	public String export(@PathVariable String fileName) throws IOException {
		String docsPath = excelService.export2003(fileName, response);
		return docsPath;
	}

	@RequestMapping(value = "/export2007/{fileName}", method = RequestMethod.GET)
	public String export2007(@PathVariable String fileName) throws IOException {
		String docsPath = excelService.export2007(fileName, response);
		return docsPath;
	}

	@RequestMapping(value = "/template/{fileName}", method = RequestMethod.GET)
	public String template(@PathVariable String fileName) throws IOException {
		String docsPath = excelService.template(fileName, response);
		return docsPath;
	}

	@RequestMapping(value = "/replace/{fileName}", method = RequestMethod.GET)
	public String replace(@PathVariable String fileName) throws IOException {
		String docsPath = excelService.replace(fileName, response);
		return docsPath;
	}
}
