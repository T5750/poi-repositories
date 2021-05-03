package t5750.poi;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.web.support.SpringBootServletInitializer;
import t5750.poi.util.Globals;

import java.io.File;

@SpringBootApplication
public class PoiApplication extends SpringBootServletInitializer {
	public static void main(String[] args) {
		File file = new File(Globals.DOC_PATH);
		if (!file.exists()) {
			file.mkdirs();
		}
		SpringApplication.run(PoiApplication.class, args);
	}
}