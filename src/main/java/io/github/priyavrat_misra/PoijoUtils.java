package io.github.priyavrat_misra;

import io.github.priyavrat_misra.annotations.Sequence;
import java.lang.reflect.Field;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoijoUtils {
  public static <T> Workbook toWorkbook(T object) throws NoSuchFieldException {
    Workbook workbook = new XSSFWorkbook();
    Class<?> workbookClass = object.getClass();
    if (workbookClass.isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      prepareSheets(workbookClass, workbook);
    }
    return workbook;
  }

  private static void prepareSheets(Class<?> workbookClass, Workbook workbook)
      throws NoSuchFieldException {
    final List<String> sheetFieldNames = getEligibleSheetFieldNames(workbookClass);
    for (String sheetFieldName : sheetFieldNames) {
      final io.github.priyavrat_misra.annotations.Sheet sheet =
          workbookClass
              .getDeclaredField(sheetFieldName)
              .getDeclaredAnnotation(io.github.priyavrat_misra.annotations.Sheet.class);
      if (sheet != null && !sheet.name().isEmpty()) {
        workbook.createSheet(WorkbookUtil.createSafeSheetName(sheet.name()));
      } else {
        // TODO: split by camel case, capitalize first word
      }
    }
  }

  /**
   * A field is eligible to be a {@link Sheet} if it is public and a {@link Collection}.
   *
   * @param workbookClass workbook's class
   * @return list of eligible field names as string.
   */
  private static List<String> getEligibleSheetFieldNames(Class<?> workbookClass) {
    final List<Field> declaredFields = Arrays.asList(workbookClass.getDeclaredFields());
    final List<String> eligibleFieldNames =
        Arrays.stream(workbookClass.getFields())
            .filter(declaredFields::contains)
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
