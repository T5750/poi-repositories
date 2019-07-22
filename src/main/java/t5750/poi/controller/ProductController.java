package t5750.poi.controller;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.validation.Valid;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.validation.BindingResult;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import t5750.poi.command.ProductForm;
import t5750.poi.converter.ProductToProductForm;
import t5750.poi.domain.Product;
import t5750.poi.service.ProductService;

@Controller
public class ProductController {
	private ProductService productService;
	private ProductToProductForm productToProductForm;

	@Autowired
	public void setProductToProductForm(
			ProductToProductForm productToProductForm) {
		this.productToProductForm = productToProductForm;
	}

	@Autowired
	public void setProductService(ProductService productService) {
		this.productService = productService;
	}

	@RequestMapping("/")
	public String redirToList() {
		return "redirect:/product/list";
	}

	@RequestMapping({ "/product/list", "/product" })
	public String listProducts(Model model) {
		model.addAttribute("products", productService.listAll());
		return "product/list";
	}

	@RequestMapping("/product/show/{id}")
	public String getProduct(@PathVariable String id, Model model) {
		model.addAttribute("product", productService.getById(Long.valueOf(id)));
		return "product/show";
	}

	@RequestMapping("product/edit/{id}")
	public String edit(@PathVariable String id, Model model) {
		Product product = productService.getById(Long.valueOf(id));
		ProductForm productForm = productToProductForm.convert(product);
		model.addAttribute("productForm", productForm);
		return "product/productform";
	}

	@RequestMapping("/product/new")
	public String newProduct(Model model) {
		model.addAttribute("productForm", new ProductForm());
		return "product/productform";
	}

	@RequestMapping(value = "/product", method = RequestMethod.POST)
	public String saveOrUpdateProduct(@Valid ProductForm productForm,
			BindingResult bindingResult) {
		if (bindingResult.hasErrors()) {
			return "product/productform";
		}
		Product savedProduct = productService
				.saveOrUpdateProductForm(productForm);
		return "redirect:/product/show/" + savedProduct.getId();
	}

	@RequestMapping("/product/delete/{id}")
	public String delete(@PathVariable String id) {
		productService.delete(Long.valueOf(id));
		return "redirect:/product/list";
	}

	@RequestMapping("/product/export")
	public void export(HttpServletResponse response) {
		// java.lang.IllegalStateException: getOutputStream() has already been
		// called for this response
		// https://stackoverflow.com/questions/30560360/i-am-getting-an-exception-java-lang-illegalstateexception-getoutputstream-ha
		productService.export(response);
	}

	@RequestMapping("/product/excelImport")
	public String excelImport(HttpServletRequest request) {
		productService.excelImport(request);
		return "redirect:/product/list";
	}
}
