package za.co.silversurge.dataprocessing.annotations;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Used to color a cell based on a value
 */
@Retention(RUNTIME)
@Target(FIELD)
public @interface ColorFilter {

    /**
     * The qualifier to use for the value
     * @return the qualifier to use for the value
     */
    Qualifier qualifier();

    /**
     * The field name to use for the value
     * @return the field name to use for the value
     */
    String fieldName() default "";

    /**
     * The color the cell would be if the qualifier is true
     * @return the color the cell would be if the qualifier is true
     */
    IndexedColors matchColor();

    /**
     * The color the cell would be if the qualifier is false
     * @return the color the cell would be if the qualifier is false
     */
    IndexedColors nonMatchColor() default IndexedColors.WHITE;

    /**
     * The value to compare against
     * @return the value to compare against
     */
    String value() default "";

}
