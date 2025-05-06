package io.github.priyavrat_misra.model;

import io.github.priyavrat_misra.annotations.Workbook;
import java.util.List;

@Workbook
public class WorkbookDto2 {
  public List<AllTypes> allTypes;

  public WorkbookDto2(List<AllTypes> allTypes) {
    this.allTypes = allTypes;
  }
}
