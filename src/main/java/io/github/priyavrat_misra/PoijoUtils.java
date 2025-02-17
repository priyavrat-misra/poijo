package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Column;
import io.github.priyavrat_misra.annotations.Sequence;
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

// TODO: add docs
public class PoijoUtils {
  public static <T> Workbook toWorkbook(@NonNull T object) {
    Workbook workbook = new XSSFWorkbook();
    Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      populateData(workbookClass, workbook, object);
    }
    return workbook;
  }

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
   * A field is eligible to be a {@link Sheet} if it is public and a {@link Collection}.
   *
   * @param workbookClass workbook's class
   * @return list of eligible field names as string.
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
    if (workbookClass.isAnnotationPresent(Sequence.class)) {
      return Arrays.stream(workbookClass.getAnnotation(Sequence.class).value())
          .filter(eligibleFieldNames::contains)
          .collect(Collectors.toList());
    } else {
      return eligibleFieldNames;
    }
  }

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
    if (rowClass.isAnnotationPresent(Sequence.class)) {
      return Arrays.stream(rowClass.getAnnotation(Sequence.class).value())
          .filter(eligibleColumnNames::contains)
          .collect(Collectors.toList());
    } else {
      return eligibleColumnNames;
    }
  }

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
