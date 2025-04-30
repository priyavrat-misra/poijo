package io.github.priyavrat_misra;

import static org.assertj.core.api.Assertions.*;

import io.github.priyavrat_misra.model.*;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
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

  static Map<CellPropertyType, Object> commonStyleProperties = new HashMap<>();
  static Map<CellPropertyType, Object> headerStyleProperties = new HashMap<>();
  static Map<CellPropertyType, Object> bodyStyleProperties = new HashMap<>();

  @BeforeAll
  static void setUp() {
    cat = new Cat("Meow", 3, "Black");

    dog = new Dog("Bark", 2, 10);

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

    commonStyleProperties.put(CellPropertyType.BORDER_LEFT, BorderStyle.THIN);
    commonStyleProperties.put(
        CellPropertyType.LEFT_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    commonStyleProperties.put(CellPropertyType.BORDER_RIGHT, BorderStyle.THIN);
    commonStyleProperties.put(
        CellPropertyType.RIGHT_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    commonStyleProperties.put(CellPropertyType.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    commonStyleProperties.put(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER);

    headerStyleProperties.put(CellPropertyType.BORDER_TOP, BorderStyle.THIN);
    headerStyleProperties.put(
        CellPropertyType.TOP_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    headerStyleProperties.put(CellPropertyType.BORDER_BOTTOM, BorderStyle.THIN);
    headerStyleProperties.put(
        CellPropertyType.BOTTOM_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    headerStyleProperties.put(
        CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.GREY_25_PERCENT.getIndex());

    bodyStyleProperties.put(CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.WHITE.getIndex());
  }

  @Test
  void nullWorkbookShouldThrowNPE() {
    assertThatThrownBy(() -> Poijo.using(null))
        .as("null workbook throw NPE")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("workbook cannot be null");
  }

  @Test
  void nullObjectShouldThrowNPE() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      assertThatThrownBy(() -> Poijo.using(workbook).map(null))
          .as("null object throw NPE")
          .isInstanceOf(NullPointerException.class)
          .hasMessage("object cannot be null");
    }
  }

  @Test
  void objectNotAnnotatedWithWorkbookShouldThrowIllegalArgumentException() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      assertThatThrownBy(() -> Poijo.using(workbook).map(new Object()))
          .as("object without @Workbook throw IllegalArgumentException")
          .isInstanceOf(IllegalArgumentException.class)
          .hasMessage(
              "Passed object's class is not annotated with io.github.priyavrat_misra.annotations.Workbook");
    }
  }

  @Test
  void sheetOrderShouldMatchOrderAnnotation() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setCats(Collections.singleton(cat));
      workbookDto.setPetDogs(Collections.singletonList(dog));
      workbookDto.setStores(Collections.singletonList(store1));
      workbookDto.setUsers(Collections.singletonList(user1));

      Poijo.using(workbook).map(workbookDto);
      assertThat(workbook.getSheetAt(0).getSheetName())
          .isEqualTo(WorkbookUtil.createSafeSheetName("Sheet: Cats"));
      assertThat(workbook.getSheetAt(1).getSheetName()).isEqualTo("Pet Dogs");
      assertThat(workbook.getSheetAt(2).getSheetName()).isEqualTo("Users");
      assertThat(workbook.getSheetAt(3).getSheetName())
          .isEqualTo(WorkbookUtil.createSafeSheetName("Sheet: Stores"));
    }
  }

  @Test
  void skippedFieldInOrderAnnotationShouldNotBeMapped() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setPetDogs(Collections.singletonList(dog));

      Poijo.using(workbook).map(workbookDto);
      workbook
          .getSheetAt(0)
          .getRow(0)
          .forEach(cell -> assertThat(cell.getStringCellValue()).isNotEqualTo("Weight"));
    }
  }

  @Test
  void nestingWithoutColumnAnnotationNestedSetShouldNotBeMapped() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setCats(Collections.singleton(cat));

      Poijo.using(workbook).map(workbookDto);
      workbook
          .getSheetAt(0)
          .getRow(0)
          .forEach(cell -> assertThat(cell.getStringCellValue()).isNotEqualTo("Empty"));
    }
  }

  @Test
  void shouldSkipEmptyCollectionsAndNotCreateEmptySheets() throws IOException {
    WorkbookDto dto = new WorkbookDto(); // nothing set

    try (Workbook workbook = new XSSFWorkbook()) {
      Poijo.using(workbook).map(dto);
      assertThat(workbook.getNumberOfSheets()).isEqualTo(0);
    }
  }

  @Test
  void sheetShouldNotBeCreatedIfNoDataInRow() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setUsers(Collections.emptyList());
      Poijo.using(workbook).map(workbookDto);
      assertThat(workbook.getNumberOfSheets()).isEqualTo(0);
    }
  }

  @Test
  void shouldMapNestedObjectsWithCorrectRowCountAfterFlattening() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setUsers(Collections.singletonList(user1));
      Poijo.using(workbook).map(workbookDto);
      assertThat(workbook.getSheetAt(0).getRow(0).getPhysicalNumberOfCells()).isEqualTo(10);
    }
  }

  @Test
  void shouldFlattenNestedObjectsAndCollections() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setStores(Collections.singletonList(store1));

      Poijo.using(workbook).map(dto);

      Row header = workbook.getSheetAt(0).getRow(0);
      List<String> headers = new ArrayList<>();
      for (Cell cell : header) headers.add(cell.getStringCellValue());

      // Check for some nested fields in header
      assertThat(headers).contains("Location Details State");
      assertThat(headers).contains("Working Hours 0");
      assertThat(headers).contains("Products 0 Specs Processor");
    }
  }

  @Test
  void shouldApplyNumberFormatForAnnotatedColumns() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setUsers(Collections.singletonList(user1)); // Address.zipcode is formatted

      Poijo.using(workbook).map(dto);
      Sheet sheet = workbook.getSheetAt(0);
      Cell headerCell = sheet.getRow(0).getCell(7);
      Cell bodyCell = sheet.getRow(1).getCell(7);

      // Address Zipcode is at index 7
      assertThat(headerCell.getStringCellValue()).isEqualTo("Address Zipcode");
      // Number format index should be set (not general)
      assertThat(bodyCell.getCellStyle().getDataFormat()).isNotZero();
    }
  }

  @Test
  void shouldMapNestedObjectsWithCorrectRowCount() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setStores(Collections.singletonList(store1));
      Poijo.using(workbook).map(dto);
    }
  }

  @Test
  void applyCellStylePropertiesToHeaderShouldAffectOnlyHeader() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setUsers(Collections.singletonList(user1));
      Poijo.using(workbook)
          .map(dto)
          .applyCellStylePropertiesToHeader(
              Collections.singletonMap(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER));

      Sheet sheet = workbook.getSheetAt(0);
      Cell headerCell = sheet.getRow(0).getCell(0);
      assertThat(headerCell.getCellStyle().getAlignment())
          .as("header row should have alignment center")
          .isEqualTo(HorizontalAlignment.CENTER);

      // Body row should NOT have style applied
      Cell bodyCell = sheet.getRow(1).getCell(0);
      assertThat(bodyCell.getCellStyle().getAlignment())
          .as("body row should not have alignment center")
          .isNotEqualTo(HorizontalAlignment.CENTER);
    }
  }

  @Test
  void applyCellStylePropertiesToBodyShouldAffectOnlyBodyRows() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setUsers(Arrays.asList(user1, user2));

      Poijo.using(workbook)
          .map(dto)
          .applyCellStylePropertiesToBody(
              Collections.singletonMap(
                  CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.LIGHT_GREEN.getIndex()));

      Sheet sheet = workbook.getSheetAt(0);
      // Header row should NOT have style applied
      Cell headerCell = sheet.getRow(0).getCell(0);
      short headerColor = headerCell.getCellStyle().getFillForegroundColor();
      assertThat(headerColor).isNotEqualTo(IndexedColors.LIGHT_GREEN.getIndex());

      // Body rows should have style
      Cell bodyCell1 = sheet.getRow(1).getCell(0);
      Cell bodyCell2 = sheet.getRow(2).getCell(0);
      assertThat(bodyCell1.getCellStyle().getFillForegroundColor())
          .isEqualTo(IndexedColors.LIGHT_GREEN.getIndex());
      assertThat(bodyCell2.getCellStyle().getFillForegroundColor())
          .isEqualTo(IndexedColors.LIGHT_GREEN.getIndex());
    }
  }

  @Test
  void applyCellStylePropertiesShouldAffectAllCells() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      WorkbookDto dto = new WorkbookDto();
      dto.setStores(Collections.singletonList(store2));
      dto.setUsers(Arrays.asList(user1, user2));

      Poijo.using(workbook)
          .map(dto)
          .applyCellStyleProperties(
              Collections.singletonMap(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER));

      workbook.forEach(
          sheet ->
              sheet.forEach(
                  row ->
                      row.forEach(
                          cell ->
                              assertThat(cell.getCellStyle().getAlignment())
                                  .isEqualTo(HorizontalAlignment.CENTER))));
    }
  }

  @Test
  void emptyStyleMapsShouldNotRemoveStyles() throws IOException {
    WorkbookDto dto = new WorkbookDto();
    dto.setCats(Collections.singleton(cat));

    try (Workbook workbook = new XSSFWorkbook()) {
      Poijo.using(workbook)
          .map(dto)
          .applyCellStylePropertiesToHeader(
              Collections.singletonMap(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER))
          .applyCellStyleProperties(Collections.emptyMap());

      Sheet sheet = workbook.getSheetAt(0);
      assertThat(sheet.getRow(0).getCell(0).getCellStyle().getAlignment())
          .as("empty styles should not replace existing ones")
          .isEqualTo(HorizontalAlignment.CENTER);
    }
  }

  @Test
  void writingAWorkbookToAFileShouldWork() throws IOException {
    Files.createDirectories(Paths.get("target/output"));
    Path path = Paths.get("target/output/workbook.xlsx");
    try (Workbook workbook = new XSSFWorkbook();
        OutputStream fileOut = Files.newOutputStream(path)) {
      WorkbookDto workbookDto = new WorkbookDto();
      workbookDto.setCats(Collections.singleton(cat));
      workbookDto.setPetDogs(Collections.singletonList(dog));
      workbookDto.setStores(Arrays.asList(store1, store2));
      workbookDto.setUsers(Arrays.asList(user1, user2, user3, user4, user5));

      Poijo.using(workbook)
          .map(workbookDto)
          .applyCellStyleProperties(commonStyleProperties)
          .applyCellStylePropertiesToHeader(headerStyleProperties)
          .applyCellStylePropertiesToBody(bodyStyleProperties)
          .write(fileOut);

      assertThat(path).as("there should be an output").exists();

      Files.deleteIfExists(path);
    }
  }
}
