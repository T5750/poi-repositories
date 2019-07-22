package t5750.poi.service;

import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import t5750.poi.command.ProductForm;
import t5750.poi.domain.Product;

public interface ProductService {
	List<Product> listAll();

	Product getById(Long id);

	Product saveOrUpdate(Product product);

	void delete(Long id);

	Product saveOrUpdateProductForm(ProductForm productForm);

	List<Product> export(HttpServletResponse response);

	List excelImport(HttpServletRequest request);
}
