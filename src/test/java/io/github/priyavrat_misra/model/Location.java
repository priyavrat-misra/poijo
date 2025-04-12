package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;
import java.util.List;

@Order({"street", "details", "city", "ph"})
public class Location {
  public String street;
  public String city;

  @Column(nested = true)
  public Details details;

  @Column(name = "Phone Numbers")
  public List<String> ph;

  public Location(String street, String city, Details details, List<String> ph) {
    this.street = street;
    this.city = city;
    this.details = details;
    this.ph = ph;
  }
}
