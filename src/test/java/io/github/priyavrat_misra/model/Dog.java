package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;

@Order({"name", "age"})
public class Dog {
  @Column(name = "Responds to")
  public String name;

  public int age;

  public int weight;

  public Dog(String name, int age, int weight) {
    this.name = name;
    this.age = age;
    this.weight = weight;
  }
}
