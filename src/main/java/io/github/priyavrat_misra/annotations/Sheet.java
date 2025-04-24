package io.github.priyavrat_misra.annotations;

import io.github.priyavrat_misra.PoijoUtils;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation to specify properties for a sheet. For now, it only has support for naming sheets.
 *
 * <p>Example usage:
 *
 * <pre><code>
 * {@literal @}Workbook
 * public class NatureReserve {
 *     {@literal @}Sheet(name = "LiveAnimals")
 *     public List&lt;Animal&gt; animals;
 *     public Set&lt;Plant&gt; plantVariants;
 * }
 * </code></pre>
 *
 * <p>If name isn't provided (e.g., {@code plantVariants}), the field name is split by camel case
 * and delimited by {@link Workbook#delimiter()} and used as the sheet name. The above will result
 * in sheets with name <i>"LiveAnimals"</i> and <i>"Plant Variants"</i>.
 *
 * @author Priyavrat Misra
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Sheet {
  /**
   * Specifies the name of the sheet in the workbook. The default value is {@link PoijoUtils#EMPTY},
   * which indicates that the sheet name will be generated from the field name automatically.
   *
   * @return the name of the sheet
   */
  String name() default PoijoUtils.EMPTY;
}
