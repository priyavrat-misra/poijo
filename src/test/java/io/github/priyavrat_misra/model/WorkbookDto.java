package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Sheet;
import io.github.priyavrat_misra.annotations.Workbook;
import java.util.List;
import java.util.Set;

@Workbook
// @Order({"stores"})
public class WorkbookDto {
  @Sheet(name = "Sheet: Cats")
  public Set<Cat> cats;

  public List<Dog> petDogs;

  @Sheet(name = "Sheet: Stores")
  public List<Store> stores;

  public List<User> users;

  public void setCats(Set<Cat> cats) {
    this.cats = cats;
  }

  public void setPetDogs(List<Dog> petDogs) {
    this.petDogs = petDogs;
  }

  public void setStores(List<Store> stores) {
    this.stores = stores;
  }

  public void setUsers(List<User> users) {
    this.users = users;
  }
}
