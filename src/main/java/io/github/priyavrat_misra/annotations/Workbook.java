package io.github.priyavrat_misra.annotations;

import io.github.priyavrat_misra.Poijo;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Annotation to mark a class for processing with {@link Poijo#map(Object)} as well as specifying
 * properties. For now, it only has support for <i>delimiter</i> property which is used when
 * generating titles for columns and sheets in the workbook.
 *
 * <p>Example usage:
 *
 * <pre><code>
 * {@literal @}Workbook(delimiter = "_")
 * public class NatureReserve {
 *     {@literal @}Sheet(name = "LiveAnimals")
 *     public List&lt;Animal&gt; animals;
 *     public Set&lt;Plant&gt; plantVariants;
 * }
 *
 * public class Animal {
 *     public String animalName;
 * }
 *
 * public class Plant {
 *     {@literal @}Column(name = "Plant")
 *     public String plantType;
 * }
 * </code></pre>
 *
 * <p>In the above example, since no {@link Column#name()} was mentioned for {@code animalName}, the
 * field name is split by camel case and delimited using {@link Workbook#delimiter()}, resulting in
 * <i>"Animal_Name"</i>. The same goes for the base class where {@code plantVariants} turns into
 * <i>"Plant_Variants"</i> and used as the sheet title.
 *
 * @see Column
 * @see Sheet
 * @author Priyavrat Misra
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Workbook {
  /**
   * Specifies the delimiter to be used when generating titles for sheets, columns and nested
   * columns. The default value is {@link Poijo#SPACE}.
   *
   * @return the delimiter as {@link String}
   */
  String delimiter() default Poijo.SPACE;
}
