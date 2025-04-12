package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;

public class Address {
  public String street;
  public String city;
  public String state;

  @Column(numberFormat = "[<=99999]00000;00000-0000")
  public int zipcode;

  public Address(String street, String city, String state, int zipcode) {
    this.street = street;
    this.city = city;
    this.state = state;
    this.zipcode = zipcode;
  }
}
