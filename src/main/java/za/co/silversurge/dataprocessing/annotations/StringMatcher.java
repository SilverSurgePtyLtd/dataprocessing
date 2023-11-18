package za.co.silversurge.dataprocessing.annotations;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Used to match a string value to a color
 */
@Retention(RUNTIME)
@Target(FIELD)
public @interface StringMatcher {

    String fieldName() default "";
    IndexedColors matchColor();
    IndexedColors nonMatchColor() default IndexedColors.WHITE;

    /**
     * The value to match
     */
    String value() default "";

}
