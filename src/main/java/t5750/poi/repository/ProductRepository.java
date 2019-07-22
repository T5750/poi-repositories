package t5750.poi.repository;

import org.springframework.data.repository.CrudRepository;

import t5750.poi.domain.Product;

public interface ProductRepository extends CrudRepository<Product, Long> {
}
