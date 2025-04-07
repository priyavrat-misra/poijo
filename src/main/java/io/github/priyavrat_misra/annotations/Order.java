package io.github.priyavrat_misra.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation to specify the order of sheets or columns in the spreadsheet. The field names are
 * case-sensitive and should match the member variables in the class annotated. Omitted field names
 * aren't considered during population, and typo-ed field names are ignored.
 *
 * <p>Example usage:
 *
 * <pre><code>
 * {@literal @}Workbook
 * public class NatureReserve {
 *     public List&lt;Animal&gt; animals;
 *     public Set&lt;Plant&gt; plants;
 * }
 *
 * {@literal @}Order({"name"})
 * public class Animal {
 *     public String name;
 *     public String type;
 * }
 *
 * {@literal @}Order({"name", "type"})
 * public class Plant {
 *     public String type;
 *     public String name;
 * }
 * </code></pre>
 *
 * In the example above, {@code NatureReserve} isn't annotated {@code @Order}, which means the
 * sheets can be in any order. Note that it is not guaranteed that the order of declaration will be
 * maintained here because {@link io.github.priyavrat_misra.PoijoUtils} uses {@link
 * Class#getDeclaredFields()} to obtain the fields, which as per the docs, <q>doesn't return in any
 * particular order</q>.
 *
 * <p>In {@code Animal}, {@code type} is omitted, so it won't show up in the spreadsheet. {@code
 * Plant} sheet will have the column <i>"Name"</i> followed by <i>"Type"</i>.
 *
 * @see Class#getDeclaredFields()
 * @author Priyavrat Misra
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Order {
  /**
   * Species the order of the sheet or column fields. The field names here are case-sensitive and
   * should match the member variables in the class annotated. Omitted field names aren't considered
   * during population, and typo-ed field names are ignored.
   *
   * @return the declared order of the field names as an array of {@link String}
   */
  String[] value();
}
