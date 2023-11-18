package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Used to generate a total column in the Excel file.
 */
@Retention(RUNTIME)
@Target(FIELD)
public @interface ColumnTotal {

}
