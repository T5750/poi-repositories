package t5750.poi.service.impl;

import java.io.*;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpEntity;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpMethod;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

import t5750.poi.command.ProductForm;
import t5750.poi.converter.ProductFormToProduct;
import t5750.poi.domain.Product;
import t5750.poi.repository.ProductRepository;
import t5750.poi.service.ExcelService;
import t5750.poi.service.ProductService;
import t5750.poi.util.ExcelExportUtil;
import t5750.poi.util.Globals;

@Service
public class ProductServiceImpl implements ProductService {
	private ProductRepository productRepository;
	private ProductFormToProduct productFormToProduct;
	@Autowired
	private ExcelService excelService;
	@Autowired
	RestTemplate restTemplate;

	@Autowired
	public ProductServiceImpl(ProductRepository productRepository,
			ProductFormToProduct productFormToProduct) {
		this.productRepository = productRepository;
		this.productFormToProduct = productFormToProduct;
	}

	@Override
	public List<Product> listAll() {
		List<Product> products = new ArrayList<>();
		productRepository.findAll().forEach(products::add); // fun with Java 8
		return products;
	}

	@Override
	public Product getById(Long id) {
		return productRepository.findOne(id);
	}

	@Override
	public Product saveOrUpdate(Product product) {
		productRepository.save(product);
		return product;
	}

	@Override
	public void delete(Long id) {
		productRepository.delete(id);
	}

	@Override
	public Product saveOrUpdateProductForm(ProductForm productForm) {
		Product savedProduct = saveOrUpdate(
				productFormToProduct.convert(productForm));
		System.out.println("Saved Product Id: " + savedProduct.getId());
		return savedProduct;
	}

	@Override
	public List<Product> export(HttpServletResponse response) {
		List<Product> products = listAll();
		ExcelExportUtil<Product> ex = new ExcelExportUtil<Product>();
		String[] headers = { "Id", "Description", "Price", "Image URL" };
		String docsPath = Globals.DOC_PATH + File.separator
				+ Globals.EXPORT_PRODUCT;
		try {
			OutputStream out = new FileOutputStream(docsPath);
			ex.exportExcel(headers, products, out);
			out.close();
			excelService.download(Globals.EXPORT_PRODUCT, docsPath, response);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return products;
	}

	@Override
	public List excelImport(HttpServletRequest request) {
		HttpHeaders headers = new HttpHeaders();
		headers.setAccept(Arrays.asList(MediaType.APPLICATION_JSON));
		HttpEntity<String> entity = new HttpEntity<String>(headers);
		List<List<Object>> list = restTemplate.exchange(
				Globals.getBasePath(request) + "excel/read/productImport",
				HttpMethod.GET, entity, List.class).getBody();
		for (int i = 1; i < list.size(); i++) {
			List<Object> objList = list.get(i);
			Long id = Long.valueOf(objList.get(0).toString());
			boolean flag = productRepository.exists(id);
			if (!flag) {
				Product product = new Product();
				product.setId(id);
				if (null != objList.get(1)) {
					product.setDescription(objList.get(1).toString());
				}
				if (objList.size() > 2 && null != objList.get(2)) {
					product.setPrice(new BigDecimal(objList.get(2).toString()));
				}
				if (objList.size() > 3 && null != objList.get(3)) {
					product.setImageUrl(objList.get(3).toString());
				}
				saveOrUpdate(product);
			} else {
			}
		}
		return list;
	}
}
