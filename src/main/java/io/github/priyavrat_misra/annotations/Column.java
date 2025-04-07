package io.github.priyavrat_misra.annotations;

import io.github.priyavrat_misra.PoijoUtils;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation to specify properties for a column.
 *
 * <p>Example usage:
 *
 * <pre><code>
 * public class Author {
 *     public String name;
 * }
 *
 * public class Book {
 *     public String titleOfTheBook;
 *
 *     {@literal @}Column(nested = true)
 *     public Author author;
 *
 *     {@literal @}Column(numberFormat = "[$$-409]#,##0;-[$$-409]#,##0")
 *     public double price;
 *
 *     {@literal @}Column(name = "Date of Publication", numberFormat = "dd/MM/yyyy")
 *     public LocalDate publicationDate;
 * }
 * </code></pre>
 *
 * <p>If name isn't provided (e.g., {@code title}), the field name is split by camel case and
 * delimited by {@link Workbook#delimiter()} and used as the column name. The above will result in
 * columns with name <i>"Title Of The Book"</i>, <i>"Author Name"</i>, <i>"Price"</i> and <i>"Date
 * of Publication"</i>. Note that the title for a nested field is the path of field names to it from
 * the base class delimited by {@link Workbook#delimiter()}.
 *
 * @author Priyavrat Misra
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Column {
  /**
   * Specifies the name of the column. The default value is {@link PoijoUtils#EMPTY}, which
   * indicates that the column name will be generated from the field name automatically.
   *
   * @return the name of the column
   */
  String name() default PoijoUtils.EMPTY;

  /**
   * Specifies the number format using which the column elements are to be formatted. Useful for
   * formatting dates, currencies, percentages, zip codes, phone numbers, etc.
   *
   * @return the number format as a {@link String}
   * @see <a href="https://www.ablebits.com/office-addins-blog/custom-excel-number-format/"
   *     target="_blank">Custom Excel number format</a>
   */
  String numberFormat() default PoijoUtils.EMPTY;

  /**
   * An indicator flag to allow {@link PoijoUtils} to traverse nested objects.
   *
   * @return {@code true}, if annotated with {@code true}. Otherwise, {@code false}.
   */
  boolean nested() default false;
}
