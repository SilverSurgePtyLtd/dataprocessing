package za.co.silversurge.dataprocessing;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFHyperlink;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import za.co.silversurge.dataprocessing.annotations.*;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.EOFException;
import java.io.IOException;
import java.lang.annotation.AnnotationFormatError;
import java.lang.reflect.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Logger;

import static za.co.silversurge.dataprocessing.ExcelUtils.createCellStyle;
import static za.co.silversurge.dataprocessing.ExcelUtils.hasAnnotation;

public abstract class ExcelFile<T> {
    private static final Logger LOGGER = Logger.getLogger(ExcelFile.class.getName());
    private Class<T> entityClass;
    private final ArrayList<T> table = new ArrayList<>();
    private final ArrayList<String> errors = new ArrayList<>();
    private final List<Integer> totalColumns = new ArrayList<>();
    protected XSSFWorkbook workbook;
    protected int startRow = 0;
    protected int startColumn = 0;
    protected int endOffset = 0;
    private final Map<Integer, ExcelColumn> sizeMap = new HashMap<>();//Todo make this to auto resize map
    private CellStyle defaultStyle;
    private final Set<ExcelOnRowCreateListener> onRowCreateListenerList = new HashSet<>();

    public ExcelFile(byte[] file) throws IOException {
        workbook = new XSSFWorkbook(new ByteArrayInputStream(file));
    }

    public ExcelFile(List<T> data, Class<T> entityClass){
        this.entityClass = entityClass;
        this.table.addAll(data);
    }

    public void addRowCreateListener(ExcelOnRowCreateListener excelOnRowCreateListener){
        this.onRowCreateListenerList.add(excelOnRowCreateListener);
    }

    private void populateEntitySheetMap(HashMap<String, List<T>> map) throws IllegalAccessException {
        for (T entity : table) {
            var fields = entityClass.getDeclaredFields();
            for (var field : fields) {
                var accessible = field.canAccess(entity);
                field.setAccessible(true);
                var sheetFilter = field.getAnnotation(SheetFilter.class);
                if(sheetFilter != null){
                    var filter = field.get(entity);
                    filter = filter == null || String.valueOf(filter).isEmpty() ? sheetFilter.nullValue() : filter;
                    var entities = map.getOrDefault((String)filter, new ArrayList<>());
                    entities.add(entity);
                    map.put(String.valueOf(filter), entities);
                }
                field.setAccessible(accessible);
            }
        }
    }


    public byte[] export() throws IOException {
        var excelCompatible = entityClass.getAnnotation(ExcelCompatible.class);
        if(excelCompatible == null){
            throw new AnnotationFormatError("Class must be excel compatible!");
        }
        ByteArrayOutputStream byteArrayInputStream = new ByteArrayOutputStream();
        workbook = new XSSFWorkbook();
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        Font defaultFont = workbook.createFont();
        defaultFont.setFontHeightInPoints((short)10);
        defaultStyle = workbook.createCellStyle();
        defaultStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        defaultStyle.setFillForegroundColor(IndexedColors.WHITE.index);
        defaultStyle.setFont(defaultFont);
        font.setBold(true);
        font.setFontHeightInPoints((short)14);
        font.setColor(IndexedColors.WHITE.index);

        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.LIGHT_BLUE.index);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setFont(font);
        var entitySheetMap = new HashMap<String, List<T>>();

        if(excelCompatible.splitSheets()){
            try {
                populateEntitySheetMap(entitySheetMap);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }else{
            entitySheetMap.put("Sheet 1", table);
        }

        for (var sheetName : entitySheetMap.keySet()) {
            int rowIndex = 0;
            XSSFSheet sheet = workbook.createSheet(sheetName);
            Row row = sheet.createRow(rowIndex++);

            sheet.createFreezePane(0, 1);

            var aiRowIndex = new AtomicInteger(rowIndex);

            fieldToHeaders(row, style, entityClass, 0, new ArrayList<>(List.of(entityClass)));

            var table = entitySheetMap.get(sheetName);

            rowIndex = aiRowIndex.get();
            try {
                int idx = 0;
                for (T entity : table) {
                    var lists = getLists(entity);
                    Row r = sheet.createRow(rowIndex++);
                    int oldColIndex = entityToExcel(entity, r, new ArrayList<>(List.of(entityClass)), 0);
                    int maxSize = getMaxList(lists);
                    for (int j = 0; j < maxSize; j++) {
                        int colindex = oldColIndex;
                        for (List<Object> list : lists) {
                            var e = list.size() > j ? list.get(j) : null;
                            if (e != null) {
                                colindex = entityToExcel(e, r, new ArrayList<>(List.of(entityClass)), colindex);
                            }
                        }
                        r = sheet.createRow(rowIndex++);
                    }
                    idx ++;
                    for (var listener : onRowCreateListenerList) {
                        listener.onRowCreate(idx, table.size(), r);
                    }
                }
            } catch (IllegalAccessException e) {
                LOGGER.severe(e.getMessage());
            }
            final var totalStyle = createTotalStyle(workbook);

            for (var colIndex : totalColumns) {
                Row totalRow = sheet.createRow(rowIndex);
                final var cell = totalRow.createCell(colIndex);
                String columnReference = CellReference.convertNumToColString(colIndex);
                cell.setCellFormula("SUM(" + columnReference + "2:" + columnReference + (rowIndex) + ")");
                cell.setCellStyle(totalStyle);
            }
            for (int i = 0; i <= row.getLastCellNum() + 5; i++) {
                autosizeColumn(sheet, i);
            }

            row.setHeightInPoints((short)20);

            for (var key : sizeMap.keySet()) {
                var column = sizeMap.get(key);
                var width = Math.round((column.width()* Units.DEFAULT_CHARACTER_WIDTH+5f)/Units.DEFAULT_CHARACTER_WIDTH*256f);
                sheet.setColumnWidth(key, width);
            }
        }

        try {
            workbook.write(byteArrayInputStream);
        } catch (IOException e) {
            LOGGER.severe(e.getMessage());
        }
        byte[] data = byteArrayInputStream.toByteArray();
        byteArrayInputStream.close();
        return data;
    }

    private CellStyle createTotalStyle(XSSFWorkbook workbook) {
        final var totalStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setBold(true); // Make the font bold
        totalStyle.setFont(font);

        totalStyle.setAlignment(HorizontalAlignment.CENTER); // Center-align the text in the cell
        totalStyle.setVerticalAlignment(VerticalAlignment.CENTER); // Center-align vertically

// Add a border around the cell
        totalStyle.setBorderTop(BorderStyle.THIN);
        totalStyle.setBorderBottom(BorderStyle.THIN);
        totalStyle.setBorderLeft(BorderStyle.THIN);
        totalStyle.setBorderRight(BorderStyle.THIN);

// Set a background color for the cell
        totalStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        totalStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        DataFormat dataFormat = workbook.createDataFormat();
        String prefix = "R";
        String formatString = "\"" + prefix + "\"#,##0.00_);[Red](\"" + prefix + "\"#,##0.00)";

        short currencyFormat = dataFormat.getFormat(formatString);
        totalStyle.setDataFormat(currencyFormat);
        return totalStyle;
    }

    protected void autosizeColumn(Sheet sheet, Integer i) {
        if(!sizeMap.containsKey(i)) sheet.autoSizeColumn(i);
    }

    protected int getMaxList(List<List<Object>> lists){
        int max = 0;
        for (var list : lists) {
            if(list.size() > max){
                max = list.size();
            }
        }
        return max;
    }

    protected List<List<Object>> getLists(T entity) throws IllegalAccessException {
        var fields = entity.getClass().getDeclaredFields();
        var lists = new ArrayList<List<Object>>();
        for (Field field : fields) {
            var accessible = field.canAccess(entity);
            field.setAccessible(true);
            if (hasAnnotation(field, ExcelColumn.class) && field.getAnnotation(ExcelColumn.class).isList()) {
                lists.add((List<Object>) field.get(entity));
            }
            field.setAccessible(accessible);
        }
        return lists;
    }

    protected Cell createCell(Row row, CellStyle style, String title, int colIndex, boolean autosize) {
        Cell c = row.createCell(colIndex++);
        c.setCellValue(title);
        c.setCellStyle(style);
        if(autosize) {
            row.getSheet().autoSizeColumn(colIndex - 1);
        }
        return c;
    }

    protected int fieldToHeaders(Row row, CellStyle style, Class cls, int colIndex, ArrayList<Class> stack){
        List<Field> fields = new ArrayList<>();
        if(cls.getSuperclass() != null){
            fields.addAll(List.of(cls.getSuperclass().getDeclaredFields()));
        }
        fields.addAll(List.of(cls.getDeclaredFields()));
        for (Field field : fields) {
            if(!containsClass(field.getType(), stack)) {
                boolean accessible = field.isAccessible();
                field.setAccessible(true);
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                ColumnTotal total = field.getAnnotation(ColumnTotal.class);
                if(total != null && !totalColumns.contains(colIndex)){
                    totalColumns.add(colIndex);
                }
                if (column != null) {
                    if ((column.recurse() || hasAnnotation(field.getType(), ExcelCompatible.class)) || (column.isList() && hasAnnotation(field.getGenericType().getClass(), ExcelCompatible.class)) && !stack.contains(field.getType())) {
                        Class<ExcelCompatible> listClass = null;
                        if(column.isList()) {
                            ParameterizedType listType = (ParameterizedType) field.getGenericType();
                            listClass = (Class<ExcelCompatible>) listType.getActualTypeArguments()[0];
                        }
                        stack.add(column.isList() ? listClass : field.getType());
                        colIndex = fieldToHeaders(row, style, column.isList() ? listClass : field.getType(), colIndex, stack);
                    } else if (!stack.contains(field.getType())) {
                        Cell c = row.createCell(colIndex++);
                        c.setCellValue(column.title());
                        c.setCellStyle(style);
                        autosizeColumn(row.getSheet(), colIndex - 1);
                    }
                }else if(column != null && column.isList()){
                    stack.add(field.getType());
                }
                field.setAccessible(accessible);
            }
        }
        return colIndex;
    }

    private int entityToExcel(Object entity, Row row, ArrayList<Class> stack, int colIndex) throws IllegalAccessException {
        List<Field> fields = new ArrayList<>();
        if(entity.getClass().getSuperclass() != null){
            fields.addAll(List.of(entity.getClass().getSuperclass().getDeclaredFields()));
        }
        fields.addAll(List.of(entity.getClass().getDeclaredFields()));
        for (Field field : fields) {
            if(!containsClass(field.getType(), stack)) {
                boolean accessible = field.isAccessible();
                field.setAccessible(true);
                ExcelColumn column = field.getAnnotation(ExcelColumn.class);
                if (column != null && !column.isList()) {
                    Object value;

                    if(field.getAnnotation(Formula.class) != null){
                        var getName = "get" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
                        var isName = "is" + field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);

                        Method method;
                        try {
                            method = entity.getClass().getDeclaredMethod(getName);
                            value = method.invoke(entity);
                        } catch (NoSuchMethodException e) {
                            try {
                                method = entity.getClass().getDeclaredMethod(isName);
                                value = method.invoke(entity);
                            } catch (NoSuchMethodException | InvocationTargetException ex) {
                                ex.printStackTrace();
                                value = field.get(entity);
                            }
                        } catch (InvocationTargetException e) {
                            e.printStackTrace();
                            value = field.get(entity);
                        }
                    }else {
                        value = field.get(entity);
                    }
                    if (value == null && hasAnnotation(field.getType(), ExcelCompatible.class)) {
                        colIndex = insertNull(field.getType(), row, colIndex, stack);
                    } else {
                        colIndex = fieldToEntry(row, column, entity, field, (value == null) ? "N/A" : value, colIndex, stack);
                    }

                }
                field.setAccessible(accessible);
            }
        }
        return colIndex;
    }

    private int insertNull(Class type, Row row, int colIndex, ArrayList<Class> stack){
        boolean inStack = containsClass(type, stack);
        if(hasAnnotation(type, ExcelCompatible.class) && !inStack){
            ArrayList<Class> localStack = new ArrayList<>(stack);
            localStack.add(type);
            Field[] fields = type.getDeclaredFields();
            for (Field field :
                    fields) {
                if(field.getAnnotation(ExcelColumn.class) != null || hasAnnotation(field.getType(), ExcelCompatible.class)) {
                    boolean accessible = field.canAccess(row);
                    field.setAccessible(true);
                    colIndex = insertNull(field.getType(), row, colIndex, stack);
                    field.setAccessible(accessible);
                }
            }
        }else if(!inStack){
            Cell c = row.createCell(colIndex++);
            c.setCellValue("N/A");
            if(hasAnnotation(type, Style.class)) {
                c.setCellStyle(createCellStyle(row, (Style) type.getAnnotation(Style.class)));
            }
        }
        return colIndex;
    }

    private boolean containsClass(Class cls, ArrayList<Class> stack){
        for (Class c : stack) {
            if(cls.toString().equals(c.toString())){
                return true;
            }
        }
        return false;
    }

    private int fieldToEntry(Row row, ExcelColumn column, Object obj, Field field, Object value, int colIndex, ArrayList<Class> stack) throws IllegalAccessException {
        boolean inStack = containsClass(value.getClass(), stack);
        if(value.getClass().getAnnotation(ExcelCompatible.class) != null && !inStack){
            ArrayList<Class> localStack = new ArrayList<>(stack);
            localStack.add(value.getClass());
            colIndex = entityToExcel(value, row, localStack, colIndex);
        }else if(!inStack){
            Cell c = row.createCell(colIndex++);
            ExcelIntegerValueMap integerValueMap = field.getAnnotation(ExcelIntegerValueMap.class);
            BooleanValueMap booleanValueMap = field.getAnnotation(BooleanValueMap.class);
            ExcelLink excelLink = field.getAnnotation(ExcelLink.class);
            Style styleAno = field.getAnnotation(Style.class);
            ColorFilter colorFilter = field.getAnnotation(ColorFilter.class);
            StringMatcher matcher = field.getAnnotation(StringMatcher.class);
            CurrencyFormatter currencyFormatter = field.getAnnotation(CurrencyFormatter.class);
            CellStyle style;

            if(column.width() != -1){
                sizeMap.put(colIndex - 1, column);
            }

            if(styleAno != null) {
                style = createCellStyle(row, styleAno);
            }else{
                style = workbook.createCellStyle();
                style.cloneStyleFrom(defaultStyle);
            }
            if(excelLink != null){
                final var url = excelLink.prefix() + value + excelLink.suffix();
                CellStyle hlinkstyle = workbook.createCellStyle();
                XSSFFont hlinkfont = workbook.createFont();
                hlinkfont.setUnderline(XSSFFont.U_SINGLE);
                hlinkstyle.setFont(hlinkfont);
                hlinkstyle.setFillForegroundColor(IndexedColors.BLUE.index);

                var creationHelper = workbook.getCreationHelper();
                var link = (XSSFHyperlink)creationHelper.createHyperlink(HyperlinkType.EMAIL);//TODO add link types to annotation
                link.setAddress(url);

                String sVal = String.valueOf(value);
                int startIndex = sVal.indexOf(":") + 1;
                if(startIndex == 0){
                    c.setCellValue(sVal);
                }else{
                    int endIndex = sVal.indexOf("?");
                    c.setCellValue(sVal.substring(startIndex, endIndex == -1 ? sVal.length() : endIndex));
                }
                c.setHyperlink(link);
                c.setCellStyle(hlinkstyle);
            }else if(integerValueMap != null){
                int t = Integer.parseInt(String.valueOf(value));
                var list = Arrays.stream(integerValueMap.set()).toList();
                if(t < list.size()) {
                    c.setCellValue(list.get(t));
                }else{
                    System.err.println("T was larger than the value map... : " + obj.getClass().getName());
                }
            } else if(booleanValueMap != null && value.getClass().isAssignableFrom(Boolean.class)){
                if((Boolean)value){
                    style.setFillForegroundColor(booleanValueMap.matchColor().index);
                    c.setCellValue(booleanValueMap.trueValue());
                }else{
                    style.setFillForegroundColor(booleanValueMap.nonMatchColor().index);
                    c.setCellValue(booleanValueMap.falseValue());
                }
            } else if (!column.format().isEmpty() && value instanceof Date) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(column.format());
                c.setCellValue(simpleDateFormat.format((Date) value));
            } else if(currencyFormatter != null){
                if(value.equals("N/A")) {
                    c.setCellValue("N/A");
                }else{
                    DataFormat dataFormat = workbook.createDataFormat();
                    String prefix = currencyFormatter.symbol();
                    String formatString = "\"" + prefix + "\"#,##0.00_);[Red](\"" + prefix + "\"#,##0.00)";

                    short currencyFormat = dataFormat.getFormat(formatString);
                    style.setDataFormat(currencyFormat);

                    double numericValue = Double.parseDouble(String.valueOf(value));
                    c.setCellValue(numericValue); // Set the numeric value
                }
            }else {
                if(column.format().isEmpty()){
                    c.setCellValue(String.valueOf(value));
                }else {
                    c.setCellValue(String.format(column.format(), value));
                }
            }
            if(matcher != null){
                String valueStr = String.valueOf(value);
                String compare = matcher.value();
                if(!matcher.fieldName().isEmpty()){
                    try {
                        Field f = obj.getClass().getDeclaredField(matcher.fieldName());
                        boolean accessible = f.canAccess(obj);
                        f.setAccessible(true);
                        var fieldValue = String.valueOf(f.get(obj));
                        if(fieldValue.equals(compare)){
                            style.setFillForegroundColor(matcher.matchColor().index);
                        }else{
                            style.setFillForegroundColor(matcher.nonMatchColor().index);
                        }
                        f.setAccessible(accessible);
                    } catch (NoSuchFieldException ignore) {

                    }
                }else{
                    if(valueStr.equals(compare)){
                        style.setFillForegroundColor(matcher.matchColor().index);
                    }else{
                        style.setFillForegroundColor(matcher.nonMatchColor().index);
                    }
                }
            }
            if(colorFilter != null){
                try {
                    String valueStr = String.valueOf(value);
                    valueStr = valueStr.equals("N/A") ? "0" : valueStr;
                    String compare = colorFilter.value();
                    if(!colorFilter.fieldName().isEmpty()){
                        try {
                            Field f = obj.getClass().getDeclaredField(colorFilter.fieldName());
                            boolean accessible = f.canAccess(obj);
                            f.setAccessible(true);
                            valueStr = String.valueOf(f.get(obj));
                            f.setAccessible(accessible);
                        } catch (NoSuchFieldException ignore) {

                        }
                    }
                    try {
                        var valueField = obj.getClass().getDeclaredField(compare);
                        boolean accessible = valueField.canAccess(obj);
                        valueField.setAccessible(true);
                        compare = String.valueOf(valueField.get(obj));
                        valueField.setAccessible(accessible);
                    } catch (NoSuchFieldException ignore) {
                    }
                    compare = (compare == null) ? "" : compare;
                    double a = Double.parseDouble(valueStr);
                    double b = (!compare.isEmpty()) ? Double.parseDouble(compare) : 0;
                    boolean match = false;
                    switch (colorFilter.qualifier()){
                        case LESS_THAN:
                            match = a < b;
                        break;
                        case LESS_EQUAL:
                            match = a <= b;
                            break;
                        case EQUAL:
                            if(value instanceof String){
                                match = compare.equals(valueStr);
                            }else{
                                match = a == b;
                            }
                            break;
                        case MORE_EQUAL:
                            match = a >= b;
                            break;
                        case MORE:
                            match = a > b;
                            break;
                    }
                    if(match){
                        style.setFillForegroundColor(colorFilter.matchColor().index);
                    }else{
                        style.setFillForegroundColor(colorFilter.nonMatchColor().index);
                    }
                } catch (NumberFormatException ignore) {

                }
            }
            c.setCellStyle(style);
        }
        return colIndex;
    }

    protected Row getHeaders(){
        XSSFSheet sheet= workbook.getSheetAt(0);
        return sheet.getRow(startRow);
    }

    int getRowCount(Sheet sheet) {
        int number = 0;
        for(int i = 0; i < sheet.getLastRowNum(); i++) {
            if(sheet.getRow(i)==null) {
                sheet.shiftRows(i+1, sheet.getLastRowNum(), -1);
                i--;
            }
            number = sheet.getLastRowNum() ;
        }
        return number;
    }

    public void process(Class<T> tClass){
        //creating a Sheet object to retrieve the object
        XSSFSheet sheet= workbook.getSheetAt(0);
        Row headers = getHeaders();
        try {
            for(int i = startRow+1; i <= getRowCount(sheet) - endOffset; i ++){
                Row row = sheet.getRow(i);
                Constructor<T> constructor = tClass.getConstructor();
                T obj = constructor.newInstance();
                var lastCell = row.getLastCellNum();
                if(lastCell == -1){
                    continue;
                }
                try {
                    for (int n = startColumn; n < row.getLastCellNum(); n++) {
                        Cell cell = headers.getCell(n);
                        if (cell == null)
                            continue;
                        if(cell.getStringCellValue().startsWith("Transaction Count")){
                            i = sheet.getLastRowNum() + 1;
                            continue;
                        }
                        String name = headers.getCell(n).getStringCellValue();
                        processCell(obj, name, row.getCell(n));
                    }
                    table.add(obj);
                }catch (IllegalStateException e){
                    System.out.println(e.getMessage());
                    addError(e.getMessage(), i + 1);
                }catch (NullPointerException nullPointerException){
                    addError("Could not find object with id", i + 1);
                }catch (EOFException eofException){
                    break;
                }
            }
        } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException | InstantiationException e) {
            LOGGER.severe(e.getMessage());
        }
    }

    protected void addError(String s, int i){
        for (var error : errors) {
            if(error.startsWith(s)){
                return;
            }
        }
        errors.add(s + " on row " + i);
    }

    protected abstract void processCell(T obj, String name, Cell cell) throws EOFException;

    public boolean validateHeaders(String ... headers){
        Row head = getHeaders();
        ArrayList<String> headerList = new ArrayList<>(Arrays.asList(headers));
        for (int i = startColumn; i < head.getLastCellNum(); i++) {
            Cell cell = head.getCell(i);
            if(cell == null)
                continue;
            //cell.setCellType(CellType.STRING);
            String val = cell.getStringCellValue();
            if(val != null){
                headerList.remove(val);
            }
        }

        return headerList.isEmpty();
    }

    public abstract boolean isValid();

    public ArrayList<T> getTable() {
        return table;
    }

    public ArrayList<String> getErrors(){
        return errors;
    }

    protected XSSFWorkbook getWorkbook() {
        return workbook;
    }
}
