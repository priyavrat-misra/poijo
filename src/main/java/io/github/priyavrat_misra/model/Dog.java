package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class Dog {
  @Column(name = "Responds to")
  public String name;

  public int age;
}
