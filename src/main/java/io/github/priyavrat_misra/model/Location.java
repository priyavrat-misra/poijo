package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Column;
import java.util.List;

import io.github.priyavrat_misra.annotations.Order;
import lombok.AllArgsConstructor;

@AllArgsConstructor
@Order({"street", "details", "city", "ph"})
public class Location {
  public String street;
  public String city;

  @Column(nested = true)
  public Details details;

  @Column(name = "Phone Numbers")
  public List<String> ph;
}
