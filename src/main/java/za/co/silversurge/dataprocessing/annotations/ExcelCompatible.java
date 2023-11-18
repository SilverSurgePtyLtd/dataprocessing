package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.Documented;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.TYPE;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Used to specify that a specific class is compatible with Excel Generation
 * This is useful when recursively generating columns
 */
@Documented
@Target(TYPE)
@Retention(RUNTIME)
public @interface ExcelCompatible {

    public boolean keepTitle() default false;
    boolean splitSheets() default false;

}
