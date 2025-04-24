package io.github.priyavrat_misra.annotations;

import java.util.Map;

public enum CellStyle {
  PROPERTIES;

  private Map<String, Object> map;

  public CellStyle setMap(Map<String, Object> map) {
    this.map = map;
    return PROPERTIES;
  }

  public Map<String, Object> getMap() {
    return map;
  }
}
