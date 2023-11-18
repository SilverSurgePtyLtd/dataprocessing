package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * Used to format a number as currency
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface CurrencyFormatter {
    String symbol() default "R";
}
