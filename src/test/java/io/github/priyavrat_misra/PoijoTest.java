package io.github.priyavrat_misra;

import static org.assertj.core.api.Assertions.*;

import io.github.priyavrat_misra.model.*;
import java.time.LocalDate;
import java.util.Arrays;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

@Tag("unit")
class PoijoTest {
  static Cat cat;
  static Dog dog;
  static Store store1, store2;
  static User user1, user2, user3, user4, user5;

  @BeforeAll
  static void setup() {
    cat = new Cat("Meow", 3, "Black");

    dog = new Dog("Bark", 2);

    store1 =
        new Store(
            "Tech Store",
            new Location(
                "Downtown",
                "NY",
                new Details("New York", "91101", "US"),
                Arrays.asList("2122222222", "1234567890")),
            Arrays.asList(
                new Product(1, "Laptop", 999.99, new Specs("Intel i7", "16GB", "512GB SSD")),
                new Product(2, "Smartphone", 799.99, new Specs("Snapdragon 888", "8GB", "128GB"))),
            Arrays.asList(
                new Employee(1, "Alice", "Manager"), new Employee(2, "Bob", "Sales Associate")),
            Arrays.asList("9:00 AM - 5:00 PM", "9:00 AM - 6:00 PM", "10:00 AM - 4:00 PM"));
    store2 = new Store("Apparel", null, null, null, null);

    user1 =
        new User(
            1,
            "John Doe",
            30,
            55000.75,
            new Address("123 Main St", "Springfield", "IL", 62701),
            LocalDate.parse("1995-07-15"),
            true);

    user2 =
        new User(
            2,
            "Jane Smith",
            28,
            62000.50,
            new Address("456 Oak Rd", "Chicago", "IL", 60601),
            LocalDate.parse("1997-03-22"),
            false);

    user3 =
        new User(
            3,
            "Alice Johnson",
            35,
            75000.99,
            new Address("789 Pine Ln", "New York", "NY", 10001),
            LocalDate.parse("1989-10-10"),
            true);

    user4 =
        new User(
            4,
            "Bob Brown",
            40,
            83000.10,
            new Address("321 Elm St", "Los Angeles", "CA", 90001),
            LocalDate.parse("1984-04-12"),
            false);

    user5 =
        new User(
            5,
            "Charlie White",
            50,
            95000.25,
            new Address("654 Maple Ave", "San Francisco", "CA", 941011234),
            LocalDate.parse("1974-09-28"),
            true);
  }

  @Test
  void nullWorkbookOrNullObjectShouldThrowNullPointerException() {
    assertThatThrownBy(() -> Poijo.map(null, null))
        .as("null workbook and null object")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("workbook and object are null");

    assertThatThrownBy(() -> Poijo.map(null, new Object()))
        .as("null workbook")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("workbook is null");

    assertThatThrownBy(() -> Poijo.map(new XSSFWorkbook(), null).close())
        .as("null object")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("object is null");
  }

  @Test
  void objectNotAnnotatedWithWorkbookShouldThrowIllegalArgumentException() {
    assertThatThrownBy(() -> Poijo.map(new XSSFWorkbook(), new Object()).close())
        .as("object not annotated with @Workbook")
        .isInstanceOf(IllegalArgumentException.class)
        .hasMessage(
            "Passed object's class is not annotated with io.github.priyavrat_misra.annotations.Workbook");
  }
}
