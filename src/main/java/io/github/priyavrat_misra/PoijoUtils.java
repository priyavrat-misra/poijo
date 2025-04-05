package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;
import lombok.NonNull;
import lombok.SneakyThrows;
import org.apache.commons.collections4.ListUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

/** A reflection based utility class serving as a façade for the Apache POI APIs. */
public class PoijoUtils {
  public static final String SPACE = " ";
  public static final String EMPTY = "";

  /**
   * Maps {@code object} to {@code workbook}.
   *
   * @param workbook a {@link Workbook} object
   * @param object an object to be mapped to {@code workbook}
   * @return {@code workbook} with the {@code object mapped}
   * @param <T> type parameter for {@code object}
   * @throws NullPointerException if {@code object} is {@code null}
   * @throws IllegalArgumentException if {@code object} is not annotated with {@link
   *     io.github.priyavrat_misra.annotations.Workbook}
   */
  public static <T> Workbook map(Workbook workbook, @NonNull T object) {
    final Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      populateWorkbook(workbookClass, workbook, object);
    } else {
      throw new IllegalArgumentException(
          "Passed object is not annotated with io.github.priyavrat_misra.annotations.Workbook");
    }
    return workbook;
  }

  /**
   * Gets the fields eligible for sheets, for each creates a sheet and populates the data.
   *
   * <p>{@link SneakyThrows} is used to reduce verbosity because {@link NoSuchFieldException} and
   * {@link IllegalAccessException} will never arise as {@link
   * PoijoUtils#getEligibleSheetFieldNames(Class)} only returns accessible fields.
   */
  @SneakyThrows
  private static <T> void populateWorkbook(Class<?> workbookClass, Workbook workbook, T object) {
    final io.github.priyavrat_misra.annotations.Workbook workbookAnnotation =
        workbookClass.getDeclaredAnnotation(io.github.priyavrat_misra.annotations.Workbook.class);
    final List<String> sheetFieldNames = getEligibleSheetFieldNames(workbookClass);
    for (String sheetFieldName : sheetFieldNames) {
      final Field sheetField = workbookClass.getDeclaredField(sheetFieldName);
      final Sheet sheet = createSheet(sheetField, workbook, sheetFieldName);
      final Object rows = sheetField.get(object);
      if (rows != null) {
        populateSheet(
            sheet,
            (Collection<?>) rows,
            EMPTY,
            0,
            workbookAnnotation,
            sheetField.getDeclaredAnnotation(Column.class));
      }
    }
  }

  /**
   * A field is eligible to be a {@link Sheet} if it is {@code public}, non-inherited and a {@link
   * Collection}. If {@link Order#value()} is non-empty, then returns the {@code
   * eligibleColumnNames} from it without disturbing the order.
   *
   * <p>Only {@code public} fields are considered because to access other kind of variables, the
   * accessibility level has to be altered via reflection, and altering or bypassing the
   * accessibility of classes, methods, or fields through reflection violates the encapsulation
   * principle.
   *
   * @param workbookClass workbook's class
   * @return list of eligible (and possibly ordered) field names as string
   * @see <a
   *     href="https://wiki.sei.cmu.edu/confluence/display/java/SEC05-J.+Do+not+use+reflection+to+increase+accessibility+of+classes%2C+methods%2C+or+fields">SEC05-J.
   *     Do not use reflection to increase accessibility of classes, methods, or fields</a>
   */
  private static List<String> getEligibleSheetFieldNames(Class<?> workbookClass) {
    final List<String> eligibleFieldNames =
        ListUtils.intersection(
                Arrays.asList(workbookClass.getFields()),
                Arrays.asList(workbookClass.getDeclaredFields()))
            .stream()
            .filter(field -> Collection.class.isAssignableFrom(field.getType()))
            .map(Field::getName)
            .collect(Collectors.toList());
    if (workbookClass.isAnnotationPresent(Order.class)) {
      return Arrays.stream(workbookClass.getAnnotation(Order.class).value())
          .filter(eligibleFieldNames::contains)
          .collect(Collectors.toList());
    } else {
      return eligibleFieldNames;
    }
  }

  /**
   * Creates a {@link Sheet} in {@code workbook}. If the {@code sheetField} is annotated with a
   * {@link io.github.priyavrat_misra.annotations.Sheet#name()}, then it is used as the {@link
   * Sheet} name. Otherwise, the field's name is split by camel case, capitalized and used as the
   * name.
   *
   * @param sheetField used to access {@link io.github.priyavrat_misra.annotations.Sheet}
   * @param workbook this is where the sheet is created
   * @param sheetFieldName field's name as a string
   * @return newly created {@link Sheet}'s reference
   * @see WorkbookUtil#createSafeSheetName(String)
   * @see StringUtils#capitalize(String)
   * @see StringUtils#splitByCharacterTypeCamelCase(String)
   */
  private static Sheet createSheet(Field sheetField, Workbook workbook, String sheetFieldName) {
    final io.github.priyavrat_misra.annotations.Sheet sheetAnnotation =
        sheetField.getDeclaredAnnotation(io.github.priyavrat_misra.annotations.Sheet.class);
    return workbook.createSheet(
        WorkbookUtil.createSafeSheetName(
            sheetAnnotation != null && !sheetAnnotation.name().isEmpty()
                ? sheetAnnotation.name()
                : prepareCapitalizedForm(sheetFieldName)));
  }

  /**
   * Returns all the public (refer {@link PoijoUtils#getEligibleSheetFieldNames(Class)}
   * documentation to know why), non-inherited and field types which are supported by {@code
   * Cell::setCellValue} inside {@code rowClass}. If {@link Order#value()} is non-empty, then
   * returns the {@code eligibleColumnNames} from it without disturbing the order.
   *
   * <p>To allow nested objects, it even considers fields annotated with {@link Column#nested()} set
   * to {@code true}.
   *
   * @param rowClass class of a sheet's element
   * @return list of eligible (and possibly ordered) sheet names as string
   * @see PoijoUtils#getEligibleSheetFieldNames(Class)
   */
  private static List<String> getEligibleColumnNames(Class<?> rowClass) {
    final List<String> eligibleColumnNames =
        ListUtils.intersection(
                Arrays.asList(rowClass.getFields()), Arrays.asList(rowClass.getDeclaredFields()))
            .stream()
            .filter(
                field ->
                    String.class.isAssignableFrom(field.getType())
                        || Integer.class.isAssignableFrom(field.getType())
                        || int.class.isAssignableFrom(field.getType())
                        || Double.class.isAssignableFrom(field.getType())
                        || double.class.isAssignableFrom(field.getType())
                        || Boolean.class.isAssignableFrom(field.getType())
                        || boolean.class.isAssignableFrom(field.getType())
                        || RichTextString.class.isAssignableFrom(field.getType())
                        || Date.class.isAssignableFrom(field.getType())
                        || LocalDate.class.isAssignableFrom(field.getType())
                        || LocalDateTime.class.isAssignableFrom(field.getType())
                        || Calendar.class.isAssignableFrom(field.getType())
                        || Collection.class.isAssignableFrom(field.getType())
                        || field.isAnnotationPresent(Column.class)
                            && field.getDeclaredAnnotation(Column.class).nested())
            .map(Field::getName)
            .collect(Collectors.toList());
    if (rowClass.isAnnotationPresent(Order.class)) {
      return Arrays.stream(rowClass.getAnnotation(Order.class).value())
          .filter(eligibleColumnNames::contains)
          .collect(Collectors.toList());
    } else {
      return eligibleColumnNames;
    }
  }

  /**
   * Maps the sheet with provided {@code rows} columnAnnotation-wise. If a columnAnnotation is
   * annotated with {@link Column#name()}, then it is used as the columnAnnotation name. Otherwise,
   * the columnAnnotation name is split by camel case, capitalized and used as the name.
   *
   * <p>If there is an object annotated with {@link Column#nested()} or it is a {@link Collection},
   * then it is recursively traversed, it's elements are flattened and represented in the sheet. The
   * resulting title for it is the path to the property from the root {@link
   * io.github.priyavrat_misra.annotations.Workbook#delimiter()} separated.
   *
   * <p>{@link SneakyThrows} is used to reduce verbosity because {@link NoSuchFieldException} will
   * never arise as {@link PoijoUtils#getEligibleColumnNames(Class)} only returns accessible fields.
   *
   * @param sheet to which the columnAnnotation names are populated
   * @param rows a collection of rows which are to be populated to the {@code sheet}
   * @param titlePath path of the title so far, used for nested properties.
   * @param columnIndex current column index
   * @param workbookAnnotation used to access workbook level properties
   * @param columnAnnotation used to access column properties
   * @return new column index
   */
  @SneakyThrows
  private static int populateSheet(
      Sheet sheet,
      Collection<?> rows,
      String titlePath,
      int columnIndex,
      final io.github.priyavrat_misra.annotations.Workbook workbookAnnotation,
      Column columnAnnotation) {
    final Class<?> rowClass =
        rows.stream().filter(Objects::nonNull).findFirst().map(Object::getClass).orElse(null);
    if (rowClass != null) {
      if (String.class.isAssignableFrom(rowClass)
          || Integer.class.isAssignableFrom(rowClass)
          || int.class.isAssignableFrom(rowClass)
          || Double.class.isAssignableFrom(rowClass)
          || double.class.isAssignableFrom(rowClass)
          || Boolean.class.isAssignableFrom(rowClass)
          || boolean.class.isAssignableFrom(rowClass)
          || RichTextString.class.isAssignableFrom(rowClass)
          || Date.class.isAssignableFrom(rowClass)
          || LocalDate.class.isAssignableFrom(rowClass)
          || LocalDateTime.class.isAssignableFrom(rowClass)
          || Calendar.class.isAssignableFrom(rowClass)) {
        columnIndex = populateColumn(sheet, rows, titlePath, columnIndex, columnAnnotation);
      } else if (Collection.class.isAssignableFrom(rowClass)) {
        columnIndex =
            populateCollection(sheet, rows, titlePath, columnIndex, workbookAnnotation, rowClass);
      } else {
        columnIndex =
            populateObject(sheet, rows, titlePath, columnIndex, workbookAnnotation, rowClass);
      }
    }
    return columnIndex;
  }

  private static int populateColumn(
      Sheet sheet, Collection<?> rows, String titlePath, int columnIndex, Column columnAnnotation) {
    int rowIndex = 0;
    Cell cell = getRow(sheet, rowIndex).createCell(columnIndex);
    cell.setCellValue(titlePath);
    ++rowIndex;
    final CellStyle cellStyle = getCellStyle(sheet, columnAnnotation);
    for (Object value : rows) {
      cell = getRow(sheet, rowIndex).createCell(columnIndex);
      if (value != null) {
        if (value instanceof String) {
          cell.setCellValue((String) value);
        } else if (value instanceof Integer) {
          cell.setCellValue((Integer) value);
        } else if (value instanceof Double) {
          cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
          cell.setCellValue((Boolean) value);
        } else if (value instanceof RichTextString) {
          cell.setCellValue((RichTextString) value);
        } else if (value instanceof Date) {
          cell.setCellValue((Date) value);
        } else if (value instanceof LocalDate) {
          cell.setCellValue((LocalDate) value);
        } else if (value instanceof LocalDateTime) {
          cell.setCellValue((LocalDateTime) value);
        } else if (value instanceof Calendar) {
          cell.setCellValue((Calendar) value);
        }
        if (cellStyle != null) {
          cell.setCellStyle(cellStyle);
        }
      }
      ++rowIndex;
    }
    return columnIndex + 1;
  }

  private static int populateCollection(
      Sheet sheet,
      Collection<?> rows,
      String titlePath,
      int columnIndex,
      io.github.priyavrat_misra.annotations.Workbook workbookAnnotation,
      Class<?> rowClass) {
    final int maxSize =
        rows.stream()
            .map(value -> value != null ? (Collection<?>) value : Collections.emptyList())
            .max(Comparator.comparingInt(Collection::size))
            .orElse(Collections.emptyList())
            .size();
    for (int index = 0; index < maxSize; ++index) {
      final int fIndex = index;
      columnIndex =
          populateSheet(
              sheet,
              rows.stream()
                  .map(value -> value != null ? (Collection<?>) value : Collections.emptyList())
                  .map(ArrayList::new)
                  .map(row -> fIndex < row.size() ? row.get(fIndex) : null)
                  .collect(Collectors.toList()),
              titlePath + workbookAnnotation.delimiter() + index,
              columnIndex,
              workbookAnnotation,
              rowClass.getDeclaredAnnotation(Column.class));
    }
    return columnIndex;
  }

  private static int populateObject(
      Sheet sheet,
      Collection<?> rows,
      String titlePath,
      int columnIndex,
      io.github.priyavrat_misra.annotations.Workbook workbookAnnotation,
      Class<?> rowClass)
      throws NoSuchFieldException {
    Column columnAnnotation;
    final List<String> columnNames = getEligibleColumnNames(rowClass);
    for (String columnName : columnNames) {
      final Field field = rowClass.getDeclaredField(columnName);
      columnAnnotation = field.getDeclaredAnnotation(Column.class);
      final String newTitlePath =
          titlePath
              + (titlePath.isEmpty() ? EMPTY : workbookAnnotation.delimiter())
              + (columnAnnotation != null && !columnAnnotation.name().isEmpty()
                  ? columnAnnotation.name()
                  : prepareCapitalizedForm(columnName));
      columnIndex =
          populateSheet(
              sheet,
              rows.stream().map(row -> getValue(field, row)).collect(Collectors.toList()),
              newTitlePath,
              columnIndex,
              workbookAnnotation,
              columnAnnotation);
    }
    return columnIndex;
  }

  private static CellStyle getCellStyle(Sheet sheet, Column columnAnnotation) {
    if (columnAnnotation != null && !columnAnnotation.formatCode().isEmpty()) {
      CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
      CreationHelper createHelper = sheet.getWorkbook().getCreationHelper();
      cellStyle.setDataFormat(
          createHelper.createDataFormat().getFormat(columnAnnotation.formatCode()));
      return cellStyle;
    } else {
      return null;
    }
  }

  private static String prepareCapitalizedForm(String camelCaseForm) {
    return StringUtils.capitalize(
        StringUtils.join(StringUtils.splitByCharacterTypeCamelCase(camelCaseForm), SPACE));
  }

  @SneakyThrows
  private static Object getValue(Field field, Object obj) {
    return obj != null ? field.get(obj) : null;
  }

  private static Row getRow(Sheet sheet, int currentRow) {
    Row row = sheet.getRow(currentRow);
    return row != null ? row : sheet.createRow(currentRow);
  }
}
