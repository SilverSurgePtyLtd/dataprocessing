package za.co.silversurge.dataprocessing;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import za.co.silversurge.dataprocessing.annotations.Style;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;

/**
 * Created by Jp Silver on 2023/11/18.
 */
public class ExcelUtils {

    public static  <t extends Annotation> boolean hasAnnotation(Field field, Class<t> annotationClass){
        return field.getAnnotation(annotationClass) != null;
    }

    public static boolean hasAnnotation(Class<?> c, Class<? extends Annotation> annotation){
        Annotation[] annotations = c.getAnnotations();
        for (Annotation a : annotations) {
            if(a.getClass().isAssignableFrom(annotation)){
                return true;
            }
        }
        return false;
    }

    public static CellStyle createCellStyle(Row row, Style styleAno){
        Workbook wb = row.getSheet().getWorkbook();

        Font font = wb.createFont();
        font.setBold(styleAno.bold());
        font.setColor(styleAno.fontColor());
        font.setUnderline(styleAno.underline());
        font.setItalic(styleAno.italic());
        font.setFontHeightInPoints(styleAno.fontSize());

        CellStyle style = wb.createCellStyle();
        style.setAlignment(styleAno.horizontalAlignment());
        style.setFillBackgroundColor(styleAno.backgroundFill().index);
        style.setFillForegroundColor(styleAno.foregroundFill().index);
        style.setShrinkToFit(styleAno.shrinkToFit());
        style.setVerticalAlignment(styleAno.verticalAlignment());
        style.setWrapText(styleAno.wrapText());
        style.setFillPattern(styleAno.fillPattern());
        style.setFont(font);

        style.setBorderLeft(styleAno.leftBorder());
        style.setBorderRight(styleAno.rightBorder());

        return style;
    }

}
