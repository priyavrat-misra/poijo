package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;

public class Product {
  public int id;
  public String name;
  public double price;

  @Column(nested = true)
  public Specs specs;

  public Product(int id, String name, double price, Specs specs) {
    this.id = id;
    this.name = name;
    this.price = price;
    this.specs = specs;
  }
}
