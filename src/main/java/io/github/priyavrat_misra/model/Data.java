package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Order;
import io.github.priyavrat_misra.annotations.Sheet;
import io.github.priyavrat_misra.annotations.Workbook;
import java.util.List;
import java.util.Set;

@lombok.Data
@Workbook
@Order({"users", "stores", "cats"})
public class Data {
  @Sheet(name = "Sheet: Cats")
  public Set<Cat> cats;

  public List<Dog> petDogs;

  @Sheet(name = "Sheet: Stores")
  public List<Store> stores;

  public List<User> users;
}
