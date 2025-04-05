package io.github.priyavrat_misra.annotations;

import io.github.priyavrat_misra.PoijoUtils;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/** Has to be used in a class if it is going to be used as a workbook. */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface Workbook {
  String delimiter() default PoijoUtils.SPACE;
}
