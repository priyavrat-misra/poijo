package io.github.priyavrat_misra.model;

import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Calendar;
import java.util.Date;

import io.github.priyavrat_misra.annotations.Order;
import org.apache.poi.ss.usermodel.RichTextString;

@Order({
  "str",
  "integerObj",
  "integerPrim",
  "doubleObj",
  "doublePrim",
  "booleanObj",
  "booleanPrim",
  "date",
  "localDate",
  "localDateTime",
  "calendar",
  "richText"
})
public class AllTypes {
  // String types
  public String str;

  // Integer types
  public Integer integerObj;
  public int integerPrim;

  // Double types
  public Double doubleObj;
  public double doublePrim;

  // Boolean types
  public Boolean booleanObj;
  public boolean booleanPrim;

  // Date/Time types
  public Date date;
  public LocalDate localDate;
  public LocalDateTime localDateTime;
  public Calendar calendar;

  // RichTextString
  public RichTextString richText;

  public AllTypes(
      String str,
      Integer integerObj,
      int integerPrim,
      Double doubleObj,
      double doublePrim,
      Boolean booleanObj,
      boolean booleanPrim,
      Date date,
      LocalDate localDate,
      LocalDateTime localDateTime,
      Calendar calendar,
      RichTextString richText) {
    this.str = str;
    this.integerObj = integerObj;
    this.integerPrim = integerPrim;
    this.doubleObj = doubleObj;
    this.doublePrim = doublePrim;
    this.booleanObj = booleanObj;
    this.booleanPrim = booleanPrim;
    this.date = date;
    this.localDate = localDate;
    this.localDateTime = localDateTime;
    this.calendar = calendar;
    this.richText = richText;
  }
}
