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
import org.junit.jupiter.api.AfterAll;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

@Tag("unit")
class PoijoTest {

  private final int CAT_SHEET_ID = 0;
  private final int DOG_SHEET_ID = 1;
  private final int USER_SHEET_ID = 2;
  private final int STORE_SHEET_ID = 3;

  private static Workbook workbook;
  private static WorkbookDto workbookDto;

  @BeforeAll
  static void setUp() {
    workbookDto =
        new WorkbookDto(
            Collections.singleton(new Cat("Meow", 3, "Black")),
            Collections.singletonList(new Dog("Bark", 2, 10)),
            Arrays.asList(
                new Store(
                    "Tech Store",
                    new Location(
                        "Downtown",
                        "NY",
                        new Details("New York", "91101", "US"),
                        Arrays.asList("2122222222", "1234567890")),
                    Arrays.asList(
                        new Product(
                            1, "Laptop", 999.99, new Specs("Intel i7", "16GB", "512GB SSD")),
                        new Product(
                            2, "Smartphone", 799.99, new Specs("Snapdragon 888", "8GB", "128GB"))),
                    Arrays.asList(
                        new Employee(1, "Alice", "Manager"),
                        new Employee(2, "Bob", "Sales Associate")),
                    Arrays.asList("9:00 AM - 5:00 PM", "9:00 AM - 6:00 PM", "10:00 AM - 4:00 PM")),
                new Store("Apparel", null, null, null, null)),
            Arrays.asList(
                new User(
                    1,
                    "John Doe",
                    30,
                    55000.75,
                    new Address("123 Main St", "Springfield", "IL", 62701),
                    LocalDate.parse("1995-07-15"),
                    true),
                new User(
                    2,
                    "Jane Smith",
                    28,
                    62000.50,
                    new Address("456 Oak Rd", "Chicago", "IL", 60601),
                    LocalDate.parse("1997-03-22"),
                    false),
                new User(
                    3,
                    "Alice Johnson",
                    35,
                    75000.99,
                    new Address("789 Pine Ln", "New York", "NY", 10001),
                    LocalDate.parse("1989-10-10"),
                    true),
                new User(
                    4,
                    "Bob Brown",
                    40,
                    83000.10,
                    new Address("321 Elm St", "Los Angeles", "CA", 90001),
                    LocalDate.parse("1984-04-12"),
                    false),
                new User(
                    5,
                    "Charlie White",
                    50,
                    95000.25,
                    new Address("654 Maple Ave", "San Francisco", "CA", 941011234),
                    LocalDate.parse("1974-09-28"),
                    true)));

    Map<CellPropertyType, Object> commonStyleProperties = new HashMap<>();
    commonStyleProperties.put(CellPropertyType.BORDER_LEFT, BorderStyle.THIN);
    commonStyleProperties.put(
        CellPropertyType.LEFT_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    commonStyleProperties.put(CellPropertyType.BORDER_RIGHT, BorderStyle.THIN);
    commonStyleProperties.put(
        CellPropertyType.RIGHT_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    commonStyleProperties.put(CellPropertyType.FILL_PATTERN, FillPatternType.SOLID_FOREGROUND);
    commonStyleProperties.put(CellPropertyType.ALIGNMENT, HorizontalAlignment.CENTER);

    Map<CellPropertyType, Object> headerStyleProperties = new HashMap<>();
    headerStyleProperties.put(CellPropertyType.BORDER_TOP, BorderStyle.THIN);
    headerStyleProperties.put(
        CellPropertyType.TOP_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    headerStyleProperties.put(CellPropertyType.BORDER_BOTTOM, BorderStyle.THIN);
    headerStyleProperties.put(
        CellPropertyType.BOTTOM_BORDER_COLOR, IndexedColors.GREY_50_PERCENT.getIndex());
    headerStyleProperties.put(
        CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.GREY_25_PERCENT.getIndex());

    Map<CellPropertyType, Object> bodyStyleProperties = new HashMap<>();
    bodyStyleProperties.put(CellPropertyType.FILL_FOREGROUND_COLOR, IndexedColors.WHITE.getIndex());

    workbook =
        Poijo.using(new XSSFWorkbook())
            .map(workbookDto)
            .applyCellStyleProperties(commonStyleProperties)
            .applyCellStylePropertiesToHeader(headerStyleProperties)
            .applyCellStylePropertiesToBody(bodyStyleProperties)
            .getWorkbook();
  }

  @AfterAll
  static void tearDown() throws IOException {
    workbook.close();
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
  void sheetOrderShouldMatchOrderAnnotation() {
    assertThat(workbook.getSheetAt(CAT_SHEET_ID).getSheetName())
        .isEqualTo(WorkbookUtil.createSafeSheetName("Sheet: Cats"));
    assertThat(workbook.getSheetAt(DOG_SHEET_ID).getSheetName()).isEqualTo("Pet Dogs");
    assertThat(workbook.getSheetAt(USER_SHEET_ID).getSheetName()).isEqualTo("Users");
    assertThat(workbook.getSheetAt(STORE_SHEET_ID).getSheetName())
        .isEqualTo(WorkbookUtil.createSafeSheetName("Sheet: Stores"));
  }

  @Test
  void shouldMapNestedObjectsWithCorrectRowCount() {
    assertThat(workbook.getSheetAt(CAT_SHEET_ID).getPhysicalNumberOfRows())
        .isEqualTo(workbookDto.getCats().size() + 1);
    assertThat(workbook.getSheetAt(DOG_SHEET_ID).getPhysicalNumberOfRows())
        .isEqualTo(workbookDto.getPetDogs().size() + 1);
    assertThat(workbook.getSheetAt(USER_SHEET_ID).getPhysicalNumberOfRows())
        .isEqualTo(workbookDto.getUsers().size() + 1);
    assertThat(workbook.getSheetAt(STORE_SHEET_ID).getPhysicalNumberOfRows())
        .isEqualTo(workbookDto.getStores().size() + 1);
  }

  @Test
  void skippedFieldInOrderAnnotationShouldNotBeMapped() {
    workbook
        .getSheetAt(DOG_SHEET_ID)
        .getRow(0)
        .forEach(cell -> assertThat(cell.getStringCellValue()).isNotEqualTo("Weight"));
  }

  @Test
  void nestingWithoutColumnAnnotationNestedSetShouldNotBeMapped() {
    workbook
        .getSheetAt(CAT_SHEET_ID)
        .getRow(0)
        .forEach(cell -> assertThat(cell.getStringCellValue()).isNotEqualTo("Empty"));
  }

  @Test
  void shouldSkipEmptyCollectionsAndNotCreateEmptySheets() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      Poijo.using(workbook).map(new WorkbookDto());
      assertThat(workbook.getNumberOfSheets()).isEqualTo(0);
    }
  }

  @Test
  void sheetShouldNotBeCreatedIfNoDataInRow() throws IOException {
    try (Workbook workbook = new XSSFWorkbook()) {
      Poijo.using(workbook)
          .map(
              new WorkbookDto(
                  Collections.emptySet(),
                  Collections.emptyList(),
                  Collections.emptyList(),
                  Collections.emptyList()));
      assertThat(workbook.getNumberOfSheets()).isEqualTo(0);
    }
  }

  @Test
  void shouldMapNestedObjectsWithCorrectRowCountAfterFlattening() {
    assertThat(workbook.getSheetAt(USER_SHEET_ID).getRow(0).getPhysicalNumberOfCells())
        .isEqualTo(10);
  }

  @Test
  void shouldFlattenNestedObjectsAndCollections() {
    Row header = workbook.getSheetAt(STORE_SHEET_ID).getRow(0);
    Set<String> headers = new HashSet<>();
    for (Cell cell : header) headers.add(cell.getStringCellValue());

    // Check for some nested fields in header
    assertThat(headers).contains("Location Details State");
    assertThat(headers).contains("Working Hours 0");
    assertThat(headers).contains("Products 0 Specs Processor");
  }

  @Test
  void shouldApplyNumberFormatForAnnotatedColumns() {
    // Address.zipcode is formatted
    Sheet sheet = workbook.getSheetAt(USER_SHEET_ID);
    Cell headerCell = sheet.getRow(0).getCell(7);
    Cell bodyCell = sheet.getRow(1).getCell(7);

    // Address Zipcode is at index 7
    assertThat(headerCell.getStringCellValue()).isEqualTo("Address Zipcode");
    // Number format index should be set (not general)
    assertThat(bodyCell.getCellStyle().getDataFormat()).isNotZero();
  }

  @Test
  void applyCellStylePropertiesToHeaderShouldAffectOnlyHeader() {
    Sheet sheet = workbook.getSheetAt(0);
    Cell headerCell = sheet.getRow(0).getCell(0);
    assertThat(headerCell.getCellStyle().getFillForegroundColor())
        .as("header row should have grey color")
        .isEqualTo(IndexedColors.GREY_25_PERCENT.getIndex());

    // Body row should NOT have style applied
    Cell bodyCell = sheet.getRow(1).getCell(0);
    assertThat(bodyCell.getCellStyle().getFillForegroundColor())
        .as("body row should not have grey color")
        .isNotEqualTo(IndexedColors.GREY_25_PERCENT.getIndex());
  }

  @Test
  void applyCellStylePropertiesShouldAffectAllCells() {
    workbook.forEach(
        sheet ->
            sheet.forEach(
                row ->
                    row.forEach(
                        cell ->
                            assertThat(cell.getCellStyle().getAlignment())
                                .isEqualTo(HorizontalAlignment.CENTER))));
  }

  @Test
  void emptyStyleMapsShouldNotRemoveStyles() {
    Poijo.using(workbook).applyCellStyleProperties(Collections.emptyMap());
    assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getCellStyle().getAlignment())
        .as("empty styles should not replace existing ones")
        .isEqualTo(HorizontalAlignment.CENTER);
  }

  @Test
  void writingAWorkbookToAFileShouldWork() throws IOException {
    Files.createDirectories(Paths.get("target/output"));
    Path path = Paths.get("target/output/workbook.xlsx");
    try (OutputStream fileOut = Files.newOutputStream(path)) {
      Poijo.using(workbook).write(fileOut);
      assertThat(path).as("there should be an output").exists();
      Files.deleteIfExists(path);
    }
  }
}
