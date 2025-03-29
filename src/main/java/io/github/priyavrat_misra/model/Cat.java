package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;
import lombok.AllArgsConstructor;

@AllArgsConstructor
@Order({"age", "color"})
public class Cat {
  public String name;
  public int age;

  @Column(name = "Fur Color")
  public String color;
}
