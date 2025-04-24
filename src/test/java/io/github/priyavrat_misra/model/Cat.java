package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;

@Order({"name", "age", "color"})
public class Cat {
  public String name;
  public int age;

  @Column(name = "Fur Color")
  public String color;

  public Cat(String name, int age, String color) {
    this.name = name;
    this.age = age;
    this.color = color;
  }
}
