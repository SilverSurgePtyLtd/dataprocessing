package za.co.silversurge.dataprocessing.annotations;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by Jp Silver on 2023/11/18.
 * Used to generate a column in the Excel file.
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.METHOD})
public @interface ExcelColumn {

    /**
     * The title of the column
     * @return the title of the column
     */
    String title();

    /**
     * The date time format or Excel format to be used for the column
     * @return the name of the field to be used in the column
     */
    String format() default "";

    /**
     * The column width in the Excel file
     * @return the width of the column in the Excel file
     */
    int width() default -1;

    /**
     * If the Excel file should also generate columns for this field
     * @return true if the field should be recursed
     */
    boolean recurse() default false;

    /**
     * If the field is a list, this will generate a column for each item in the list
     * @return
     */
    boolean isList() default false;

}
