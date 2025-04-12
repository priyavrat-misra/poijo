package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import java.time.LocalDate;

public class User {
  public int id;
  public String name;
  public int age;
  public double salary;

  @Column(nested = true)
  public Address address;

  @Column(name = "Date of Birth", numberFormat = "dd/MM/yyyy")
  public LocalDate dob;

  public boolean active;

  public User(
      int id, String name, int age, double salary, Address address, LocalDate dob, boolean active) {
    this.id = id;
    this.name = name;
    this.age = age;
    this.salary = salary;
    this.address = address;
    this.dob = dob;
    this.active = active;
  }
}
