package za.co.silversurge.dataprocessing.annotations;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Used to map a boolean value to a string value
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface BooleanValueMap {

    /**
     * The value to match with true
     * @return the value to match with true
     */
    String trueValue();

    /**
     * The value to match with false
     * @return the value to match with false
     */
    String falseValue();

    /**
     * The color to use if the value is true
     * @return the color to use if the value matches
     */
    IndexedColors matchColor() default IndexedColors.WHITE;

    /**
     * The color to use if the value is false
     * @return the color to use if the value does not match
     */
    IndexedColors nonMatchColor() default IndexedColors.WHITE;

}
