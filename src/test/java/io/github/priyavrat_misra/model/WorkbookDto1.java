package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Order;
import io.github.priyavrat_misra.annotations.Sheet;
import java.util.List;
import java.util.Set;

@Order({"cats", "petDogs", "users", "stores"})
public class WorkbookDto1 {
  @Sheet(name = "Sheet: Cats")
  public Set<Cat> cats;

  public List<Dog> petDogs;

  @Sheet(name = "Sheet: Stores")
  public List<Store> stores;

  @Sheet public List<User> users;

  public Set<Cat> getCats() {
    return cats;
  }

  public List<Dog> getPetDogs() {
    return petDogs;
  }

  public List<Store> getStores() {
    return stores;
  }

  public List<User> getUsers() {
    return users;
  }

  public WorkbookDto1() {}

  public WorkbookDto1(Set<Cat> cats, List<Dog> petDogs, List<Store> stores, List<User> users) {
    this.cats = cats;
    this.petDogs = petDogs;
    this.stores = stores;
    this.users = users;
  }
}
