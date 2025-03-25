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

  private static CreationHelper createHelper;

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
    final Workbook workbook = new XSSFWorkbook();
    createHelper = workbook.getCreationHelper();
    final Class<?> workbookClass = object.getClass();
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
      final Field sheetField = workbookClass.getDeclaredField(sheetFieldName);
      final Sheet sheet = createSheet(sheetField, workbook, sheetFieldName);
      final Collection<?> rows = (Collection<?>) sheetField.get(object);
      rows.stream()
          .findFirst()
          .map(Object::getClass)
          .ifPresent(rowClass -> populateSheet(sheet, rowClass, rows, StringUtils.EMPTY, 0));
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
   * Maps the sheet with provided {@code rows} column-wise. If a column is annotated with {@link
   * Column#name()}, then it is used as the column name. Otherwise, the column name is split by
   * camel case, capitalized and used as the name.
   *
   * <p>If there is an object annotated with {@link Column#nested()}, then it is recursively
   * traversed, it's properties are flattened and represented in the sheet. The resulting title for
   * it is the path to the property from the root space-separated.
   *
   * <p>{@link SneakyThrows} is used to reduce verbosity because {@link NoSuchFieldException} will
   * never arise as {@link PoijoUtils#getEligibleColumnNames(Class)} only returns accessible fields.
   *
   * @param sheet to which the column names are populated
   * @param rowClass class of a sheet's element
   * @param rows a collection of rows which are to be populated to the {@code sheet}
   * @param path used to name nested fields
   * @param currentCol current column index
   * @return next column index after an object is mapped
   */
  @SneakyThrows
  private static int populateSheet(
      Sheet sheet, Class<?> rowClass, Collection<?> rows, String path, int currentCol) {
    final List<String> columnNames = getEligibleColumnNames(rowClass);
    for (String columnName : columnNames) {
      final Field field = rowClass.getDeclaredField(columnName);
      final Column columnAnnotation = field.getDeclaredAnnotation(Column.class);
      final String title =
          path
              + (path.isEmpty() ? StringUtils.EMPTY : StringUtils.SPACE)
              + (columnAnnotation != null && !columnAnnotation.name().isEmpty()
                  ? columnAnnotation.name()
                  : prepareCapitalizedForm(columnName));
      if (columnAnnotation != null && columnAnnotation.nested()) {
        // if it is nested, recursively populate the sheet by flattening it
        currentCol =
            populateSheet(
                sheet,
                field.getType(),
                rows.stream().map(obj -> getValue(field, obj)).collect(Collectors.toList()),
                title,
                currentCol);
      } else {
        // otherwise, populate the current column
        currentCol = populateColumn(sheet, field, rows, title, currentCol);
      }
    }
    return currentCol;
  }

  /**
   * Populates title and values to a column with index {@code columnIndex}.
   *
   * @param sheet to which the column names are populated
   * @param field used to access the value of the field
   * @param rows a collection of rows which are to be populated to the {@code sheet}
   * @param title title for the column
   * @param columnIndex current column index
   * @return next column index after an object is mapped
   */
  private static int populateColumn(
      Sheet sheet, Field field, Collection<?> rows, String title, int columnIndex) {
    // set column title
    int rowIndex = 0;
    Cell cell = getRow(sheet, rowIndex).createCell(columnIndex);
    cell.setCellValue(title);
    ++rowIndex;
    // set column values
    final CellStyle cellStyle = getCellStyle(sheet, field);
    for (Object rowObj : rows) {
      cell = getRow(sheet, rowIndex).createCell(columnIndex);
      final Object value = getValue(field, rowObj);
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

  private static CellStyle getCellStyle(Sheet sheet, Field field) {
    final Column columnAnnotation = field.getDeclaredAnnotation(Column.class);
    if (columnAnnotation != null && !columnAnnotation.formatCode().isEmpty()) {
      CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
      cellStyle.setDataFormat(
          createHelper.createDataFormat().getFormat(columnAnnotation.formatCode()));
      return cellStyle;
    } else {
      return null;
    }
  }

  private static String prepareCapitalizedForm(String camelCaseForm) {
    return StringUtils.capitalize(
        StringUtils.join(
            StringUtils.splitByCharacterTypeCamelCase(camelCaseForm), StringUtils.SPACE));
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
