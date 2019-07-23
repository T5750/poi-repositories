package t5750.poi.export;

import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import t5750.poi.PoiApplication;
import t5750.poi.service.ExcelService;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = PoiApplication.class, webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
public class TestExportExcel {
	@Autowired
	private ExcelService excelService;
	@Autowired
	private HttpServletResponse response;

	@Test
	public void exportExcel() throws IOException {
		String fileName = "default";
		String docsPath = excelService.export2003(fileName, response);
		System.out.println(docsPath);
	}
}