package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;

public class Dog {
  @Column(name = "Responds to")
  public String name;

  public Dog(String name, int age) {
    this.name = name;
    this.age = age;
  }

  public int age;
}
