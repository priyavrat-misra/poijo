package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Sequence;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import lombok.NonNull;
import lombok.SneakyThrows;
import org.apache.commons.collections4.ListUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoijoUtils {
  public static <T> Workbook toWorkbook(@NonNull T object) {
    Workbook workbook = new XSSFWorkbook();
    Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      prepareSheets(workbookClass, workbook);
    }
    return workbook;
  }

  private static void prepareSheets(Class<?> workbookClass, Workbook workbook) {
    final List<String> sheetFieldNames = getEligibleSheetFieldNames(workbookClass);
    for (String sheetFieldName : sheetFieldNames) {
      final Sheet sheet = createSheet(workbookClass, workbook, sheetFieldName);
    }
  }

  @SneakyThrows
  private static Sheet createSheet(
      Class<?> workbookClass, Workbook workbook, String sheetFieldName) {
    final io.github.priyavrat_misra.annotations.Sheet sheetAnnotation =
        workbookClass
            .getDeclaredField(sheetFieldName)
            .getDeclaredAnnotation(io.github.priyavrat_misra.annotations.Sheet.class);
    return workbook.createSheet(
        WorkbookUtil.createSafeSheetName(
            sheetAnnotation != null && !sheetAnnotation.name().isEmpty()
                ? sheetAnnotation.name()
                : StringUtils.capitalize(
                    String.join(
                        StringUtils.SPACE,
                        StringUtils.splitByCharacterTypeCamelCase(sheetFieldName)))));
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
}
