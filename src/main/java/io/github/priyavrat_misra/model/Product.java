package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class Product {
  public int id;
  public String name;
  public double price;
  @Column(nested = true)
  public Specs specs;
}
