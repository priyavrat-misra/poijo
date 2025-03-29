package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import java.time.LocalDate;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class User {
  public int id;
  public String name;
  public int age;
  public double salary;

  @Column(nested = true)
  public Address address;

  @Column(name = "Date of Birth", formatCode = "yyyy-MM-dd")
  public LocalDate dob;

  public boolean active;
}
