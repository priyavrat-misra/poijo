package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import java.util.List;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class Store {
  public String name;

  @Column(nested = true)
  public Location location;

  public List<Product> products;
  public List<Employee> employees;

  @Column(name = "Working Hours")
  public List<String> working_hours;
}
