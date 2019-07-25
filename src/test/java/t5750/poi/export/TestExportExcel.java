package t5750.poi.export;

import java.io.File;
import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.test.context.junit4.SpringRunner;

import t5750.poi.PoiApplication;
import t5750.poi.service.ExcelService;
import t5750.poi.util.ExcelExportUtil;
import t5750.poi.util.Globals;
import t5750.poi.util.TestUtil;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = PoiApplication.class, webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
public class TestExportExcel {
	@Test
	public void exportExcel() throws IOException {
		String fileName = "default";
		String docsPath;
		Resource resource = new ClassPathResource(Globals.DOC + File.separator
				+ fileName + Globals.SUFFIX_XLS);
		if (resource.exists()) {
			docsPath = resource.getFile().getPath();
		} else {
			String imagesPath = TestUtil.IMG_PATH + File.separator + "tomcat"
					+ Globals.SUFFIX_PNG;
			ExcelExportUtil.export2003(imagesPath, TestUtil.DOC_PATH);
			docsPath = TestUtil.DOC_PATH + File.separator + Globals.EXPORT_BOOK;
		}
		System.out.println(docsPath);
	}
}