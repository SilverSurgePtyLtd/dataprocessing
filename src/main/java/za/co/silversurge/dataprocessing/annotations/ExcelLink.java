package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Used to create a hyperlink in the Excel file, eg/ Websites emails etc.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelLink {

    /**
     * The prefix and suffix to be added to the value of the field
     */
    String prefix() default "";
    String suffix() default "";

}
