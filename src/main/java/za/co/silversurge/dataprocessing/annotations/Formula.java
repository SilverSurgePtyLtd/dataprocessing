package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import static java.lang.annotation.ElementType.FIELD;
import static java.lang.annotation.ElementType.METHOD;
import static java.lang.annotation.RetentionPolicy.RUNTIME;

/**
 * Used to create a formula in the Excel file, eg/ SUM(A1:A10)
 * NOT IMPLEMENTED YET
 */
@Target({METHOD, FIELD})
@Retention(RUNTIME)
public @interface Formula {
	String value();
}