package io.github.priyavrat_misra;

import static org.assertj.core.api.Assertions.*;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Tag;
import org.junit.jupiter.api.Test;

@Tag("unit")
class PoijoTest {
  @Test
  void nullWorkbookOrNullObjectShouldThrowNullPointerException() {
    assertThatThrownBy(() -> Poijo.map(null, null))
        .as("null workbook and null object")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("workbook and object are null");

    assertThatThrownBy(() -> Poijo.map(null, new Object()))
        .as("null workbook")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("workbook is null");

    assertThatThrownBy(() -> Poijo.map(new XSSFWorkbook(), null))
        .as("null object")
        .isInstanceOf(NullPointerException.class)
        .hasMessage("object is null");
  }
}
