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
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** A reflection based utility class serving as a façade for the Apache POI APIs. */
public class PoijoUtils {
  /**
   * Maps {@code object} to a {@link XSSFWorkbook} object.
   *
   * @param object to be mapped to a {@link Workbook}
   * @return a {@link XSSFWorkbook} with details populated
   * @param <T> type parameter for {@code object}
   * @throws NullPointerException if {@code object} is {@code null}
   * @throws IllegalArgumentException if {@code object} is not annotated with {@link
   *     io.github.priyavrat_misra.annotations.Workbook}
   */
  public static <T> Workbook toWorkbook(@NonNull T object) {
    Workbook workbook = new XSSFWorkbook();
    Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      populateData(workbookClass, workbook, object);
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
  private static <T> void populateData(Class<?> workbookClass, Workbook workbook, T object) {
    final List<String> sheetFieldNames = getEligibleSheetFieldNames(workbookClass);
    for (String sheetFieldName : sheetFieldNames) {
      Field sheetField = workbookClass.getDeclaredField(sheetFieldName);
      final Sheet sheet = createSheet(sheetField, workbook, sheetFieldName);
      Collection<?> rows = (Collection<?>) sheetField.get(object);
      Class<?> rowClass = rows.stream().findFirst().map(Object::getClass).orElse(null);
      if (rowClass != null) {
        final List<String> columnNames = getEligibleColumnNames(rowClass);
        populateTitle(sheet, columnNames, rowClass);
        populateBody(sheet, columnNames, rows, rowClass);
      }
    }
  }

  /**
   * A field is eligible to be a {@link Sheet} if it is {@code public}, non-inherited and a {@link
   * Collection}. If {@link Order#value()} is non-empty, then returns the {@code eligibleColumnNames}
   * from it without disturbing the order.
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
                : StringUtils.capitalize(
                    StringUtils.join(
                        StringUtils.splitByCharacterTypeCamelCase(sheetFieldName),
                        StringUtils.SPACE))));
  }

  /**
   * Returns all the public (refer {@link PoijoUtils#getEligibleSheetFieldNames(Class)}
   * documentation to know why), non-inherited and field types which are supported by {@code
   * Cell::setCellValue} inside {@code rowClass}. If {@link Order#value()} is non-empty, then
   * returns the {@code eligibleColumnNames} from it without disturbing the order.
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
                        || Double.class.isAssignableFrom(field.getType())
                        || Boolean.class.isAssignableFrom(field.getType())
                        || RichTextString.class.isAssignableFrom(field.getType())
                        || Date.class.isAssignableFrom(field.getType())
                        || LocalDate.class.isAssignableFrom(field.getType())
                        || LocalDateTime.class.isAssignableFrom(field.getType())
                        || Calendar.class.isAssignableFrom(field.getType()))
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
   * Maps the first row of the sheet with column names. If a column is annotated with {@link
   * Column#name()}, then it is used as the column name. Otherwise, the column name is split by
   * camel case, capitalized and used as the name.
   *
   * <p>{@link SneakyThrows} is used to reduce verbosity because {@link NoSuchFieldException} will
   * never arise as {@link PoijoUtils#getEligibleColumnNames(Class)} only returns accessible fields.
   *
   * @param sheet to which the column names are populated
   * @param columnNames possibly ordered sequence of column names
   * @param rowClass class of a sheet's element
   */
  @SneakyThrows
  private static void populateTitle(Sheet sheet, List<String> columnNames, Class<?> rowClass) {
    int currentRow = 0;
    int currentCol = 0;
    Row row = sheet.createRow(currentRow);
    for (String columnName : columnNames) {
      Cell cell = row.createCell(currentCol);
      final Column columnAnnotation =
          rowClass.getDeclaredField(columnName).getDeclaredAnnotation(Column.class);
      cell.setCellValue(
          columnAnnotation != null && !columnAnnotation.name().isEmpty()
              ? columnAnnotation.name()
              : StringUtils.capitalize(
                  StringUtils.join(
                      StringUtils.splitByCharacterTypeCamelCase(columnName), StringUtils.SPACE)));
      currentCol++;
    }
  }

  /**
   * Maps the remaining rows of the sheet with {@code rows}' values.
   *
   * <p>{@link SneakyThrows} is used to reduce verbosity because {@link NoSuchFieldException} and
   * {@link IllegalAccessException} will never arise as {@link
   * PoijoUtils#getEligibleColumnNames(Class)} only returns accessible fields.
   *
   * @param sheet to which the column names are populated
   * @param columnNames possibly ordered sequence of column names
   * @param rows a collection of rows which are to be populated to the {@code sheet}
   * @param rowClass class of a sheet's element
   */
  @SneakyThrows
  private static void populateBody(
      Sheet sheet, List<String> columnNames, Collection<?> rows, Class<?> rowClass) {
    int currentRow = 1;
    for (Object rowObject : rows) {
      Row row = sheet.createRow(currentRow);
      int currentCol = 0;
      for (String columnName : columnNames) {
        Cell cell = row.createCell(currentCol);
        Object value = rowClass.getDeclaredField(columnName).get(rowObject);
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
        }
        currentCol++;
      }
      currentRow++;
    }
  }
}
