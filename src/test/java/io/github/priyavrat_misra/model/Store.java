package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import java.util.List;

public class Store {
  public String name;

  @Column(nested = true)
  public Location location;

  public List<Product> products;
  public List<Employee> employees;

  @Column(name = "Working Hours")
  public List<String> working_hours;

  public Store(
      String name,
      Location location,
      List<Product> products,
      List<Employee> employees,
      List<String> working_hours) {
    this.name = name;
    this.location = location;
    this.products = products;
    this.employees = employees;
    this.working_hours = working_hours;
  }
}
