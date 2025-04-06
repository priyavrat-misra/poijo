package io.github.priyavrat_misra.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import org.apache.commons.lang3.StringUtils;

/** Can be used to specify the column's title. */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Column {
  String name() default StringUtils.EMPTY;

  String numberFormat() default StringUtils.EMPTY;

  boolean nested() default false;
}
