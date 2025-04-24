package io.github.priyavrat_misra;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.CellPropertyType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Poijo {
  private final Workbook workbook;

  private static final Logger logger = LoggerFactory.getLogger(Poijo.class);

  private Poijo(Workbook workbook) {
    this.workbook = workbook;
  }

  public static Poijo using(Workbook workbook) {
    if (workbook == null) {
      logger.error("workbook is null");
      throw new NullPointerException("workbook cannot be null");
    }
    return new Poijo(workbook);
  }

  public <T> Poijo map(T object) {
    PoijoUtils.map(workbook, validate(object));
    return this;
  }

  /**
   * @throws NullPointerException if {@code object} is {@code null}
   * @throws IllegalArgumentException if {@link T} is not annotated with {@link
   *     io.github.priyavrat_misra.annotations.Workbook}
   */
  private <T> T validate(T object) {
    if (object == null) {
      logger.error("object is null");
      throw new NullPointerException("object cannot be null");
    }
    if (!object
        .getClass()
        .isAnnotationPresent(io.github.priyavrat_misra.annotations.Workbook.class)) {
      logger.error(
          "{} is not annotated with io.github.priyavrat_misra.annotations.Workbook",
          object.getClass().getName());
      throw new IllegalArgumentException(
          "Passed object's class is not annotated with io.github.priyavrat_misra.annotations.Workbook");
    }
    return object;
  }

  public Poijo applyCellStyleProperties(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet ->
            sheet.forEach(
                row -> row.forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles))));
    return this;
  }

  public Poijo applyCellStylePropertiesToHeader(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet ->
            sheet
                .getRow(sheet.getFirstRowNum())
                .forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles)));
    return this;
  }

  public Poijo applyCellStylePropertiesToBody(Map<CellPropertyType, Object> styles) {
    workbook.forEach(
        sheet -> {
          Iterator<Row> rowIterator = sheet.rowIterator();
          // skip header row
          if (rowIterator.hasNext()) {
            rowIterator.next();
          }
          rowIterator.forEachRemaining(
              row -> row.forEach(cell -> CellUtil.setCellStylePropertiesEnum(cell, styles)));
        });
    return this;
  }

  public void write(OutputStream stream) throws IOException {
    workbook.write(stream);
  }

  public Workbook getWorkbook() {
    return workbook;
  }
}
