package t5750.poi.read;

import java.io.File;
import java.io.IOException;
import java.util.List;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;
import org.springframework.test.context.junit4.SpringRunner;

import t5750.poi.PoiApplication;
import t5750.poi.service.ExcelService;
import t5750.poi.util.Globals;

@RunWith(SpringRunner.class)
@SpringBootTest(classes = PoiApplication.class, webEnvironment = SpringBootTest.WebEnvironment.RANDOM_PORT)
public class TestReadExcel {
	@Autowired
	private ExcelService readService;

	@Test
	public void readExcel() throws IOException {
		String fileName = "testRead";
		Resource docsPath = new ClassPathResource(Globals.DOC + File.separator
				+ fileName + Globals.SUFFIX_XLS);
		List<List<Object>> list = readService.readExcel(docsPath.getFile());
	}
}
