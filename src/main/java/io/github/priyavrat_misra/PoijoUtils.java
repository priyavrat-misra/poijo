package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Order;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;
import java.util.stream.Collectors;
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
   * @throws NullPointerException if {@code workbook} or {@code object} is {@code null}
   * @throws IllegalArgumentException if {@code object} is not annotated with {@link
   *     io.github.priyavrat_misra.annotations.Workbook}
   */
  public static <T> Workbook map(Workbook workbook, T object) {
    if (workbook == null || object == null) {
      throw new NullPointerException("workbook or object is null");
    }
    final Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      populateWorkbook(workbookClass, workbook, object);
    } else {
      throw new IllegalArgumentException(
          "Passed object is not annotated with io.github.priyavrat_misra.annotations.Workbook");
    }
    return workbook;
  }

  /** Gets the fields eligible for sheets, for each creates a sheet and populates the data. */
  private static <T> void populateWorkbook(Class<?> workbookClass, Workbook workbook, T object) {
    final io.github.priyavrat_misra.annotations.Workbook workbookAnnotation =
        workbookClass.getDeclaredAnnotation(io.github.priyavrat_misra.annotations.Workbook.class);
    final List<String> sheetFieldNames = getEligibleSheetFieldNames(workbookClass);
    for (String sheetFieldName : sheetFieldNames) {
      final Field sheetField = getField(workbookClass, sheetFieldName);
      assert sheetField != null;
      final Object rows = getFieldValue(sheetField, object);
      if (rows != null) {
        final Sheet sheet = createSheet(sheetField, workbook, sheetFieldName);
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
   * Maps the sheet with provided {@code rows} column-wise. If a column is annotated with {@link
   * Column#name()}, then it is used as the column name. Otherwise, the column name is split by
   * camel case, capitalized and used as the name.
   *
   * <p>If there is an object annotated with {@link Column#nested()} or it is a {@link Collection},
   * then it is recursively traversed, it's elements are flattened and represented in the sheet. The
   * resulting title for it is the path to the property from the root {@link
   * io.github.priyavrat_misra.annotations.Workbook#delimiter()} separated.
   *
   * @param sheet to which the column names are populated
   * @param rows a collection of rows which are to be populated to the {@code sheet}
   * @param titlePath path of the title so far, used for nested properties.
   * @param columnIndex current column index
   * @param workbookAnnotation used to access workbook level properties
   * @param columnAnnotation used to access column properties
   * @return new column index
   */
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
      if (isSupportedPrimitive(rowClass)) {
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

  private static boolean isSupportedPrimitive(Class<?> clazz) {
    return String.class.isAssignableFrom(clazz)
        || Integer.class.isAssignableFrom(clazz)
        || int.class.isAssignableFrom(clazz)
        || Double.class.isAssignableFrom(clazz)
        || double.class.isAssignableFrom(clazz)
        || Boolean.class.isAssignableFrom(clazz)
        || boolean.class.isAssignableFrom(clazz)
        || RichTextString.class.isAssignableFrom(clazz)
        || Date.class.isAssignableFrom(clazz)
        || LocalDate.class.isAssignableFrom(clazz)
        || LocalDateTime.class.isAssignableFrom(clazz)
        || Calendar.class.isAssignableFrom(clazz);
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
        setCellValue(cell, value);
        if (cellStyle != null) {
          cell.setCellStyle(cellStyle);
        }
      }
      ++rowIndex;
    }
    return columnIndex + 1;
  }

  private static Row getRow(Sheet sheet, int currentRow) {
    Row row = sheet.getRow(currentRow);
    return row != null ? row : sheet.createRow(currentRow);
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

  private static void setCellValue(Cell cell, Object value) {
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
      Class<?> rowClass) {
    final List<String> columnNames = getEligibleColumnNames(rowClass);
    for (String columnName : columnNames) {
      final Field field = getField(rowClass, columnName);
      assert field != null;
      final Column columnAnnotation = field.getDeclaredAnnotation(Column.class);
      final String newTitlePath =
          titlePath
              + (titlePath.isEmpty() ? EMPTY : workbookAnnotation.delimiter())
              + (columnAnnotation != null && !columnAnnotation.name().isEmpty()
                  ? columnAnnotation.name()
                  : prepareCapitalizedForm(columnName));
      columnIndex =
          populateSheet(
              sheet,
              rows.stream().map(row -> getFieldValue(field, row)).collect(Collectors.toList()),
              newTitlePath,
              columnIndex,
              workbookAnnotation,
              columnAnnotation);
    }
    return columnIndex;
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
                    isSupportedPrimitive(field.getType())
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

  private static String prepareCapitalizedForm(String camelCaseForm) {
    return StringUtils.capitalize(
        StringUtils.join(StringUtils.splitByCharacterTypeCamelCase(camelCaseForm), SPACE));
  }

  /**
   * {@link NoSuchFieldException} is ignored because it will never arise due to {@link
   * PoijoUtils#getEligibleColumnNames(Class)} which only returns accessible fields.
   */
  private static Field getField(Class<?> clazz, String fieldName) {
    try {
      return clazz.getDeclaredField(fieldName);
    } catch (NoSuchFieldException ignored) {
      return null;
    }
  }

  /**
   * {@link IllegalAccessException} is ignored because it will never arise due to {@link
   * PoijoUtils#getEligibleColumnNames(Class)} which only returns accessible fields.
   */
  private static Object getFieldValue(Field field, Object obj) {
    try {
      return obj != null ? field.get(obj) : null;
    } catch (IllegalAccessException ignored) {
      return null;
    }
  }
}
