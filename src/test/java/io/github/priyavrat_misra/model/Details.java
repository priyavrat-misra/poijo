package io.github.priyavrat_misra.model;

public class Details {
  public String state;
  public String zip;

  public Details(String state, String zip, String country) {
    this.state = state;
    this.zip = zip;
    this.country = country;
  }

  public String country;
}
