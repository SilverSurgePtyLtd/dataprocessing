package za.co.silversurge.dataprocessing.annotations;

import org.apache.poi.ss.usermodel.*;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Used to style a cell
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Style {
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.LEFT;

    IndexedColors backgroundFill() default IndexedColors.WHITE;

    IndexedColors foregroundFill() default IndexedColors.BLACK;

    boolean shrinkToFit() default false;

    VerticalAlignment verticalAlignment() default VerticalAlignment.TOP;

    boolean wrapText() default false;

    boolean bold() default false;

    short fontColor() default 8;

    byte underline() default 0;

    boolean italic() default false;

    BorderStyle leftBorder() default BorderStyle.NONE;
    BorderStyle rightBorder() default BorderStyle.NONE;

    short fontSize() default 18;

    FillPatternType fillPattern() default FillPatternType.SOLID_FOREGROUND;
}
