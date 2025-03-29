package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import lombok.AllArgsConstructor;

@AllArgsConstructor
public class Address {
  public String street;
  public String city;
  public String state;

  @Column(formatCode = "[<=99999]00000;00000-0000")
  public int zipcode;
}
