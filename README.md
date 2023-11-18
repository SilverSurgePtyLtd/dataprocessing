# Excel Generation Library for Java

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://opensource.org/licenses/MIT)

This Java library facilitates the generation and import of Microsoft Excel files. It leverages the Apache POI library and provides a flexible and customizable way to work with Excel data.

## Features

- Generate Excel files programmatically.
- Import data from Excel files into Java objects.
- Customize Excel columns using annotations like `@ExcelColumn`, `@BooleanValueMap`, and more.

## Getting Started

### Prerequisites

- Java 8 or higher.
- Apache POI library (version X.X.X).

### Installation

Add the library to your project by including the JAR file or using a build tool like Maven or Gradle.

```xml
<!-- Maven -->
Coming Soon!
```

## Annotations

- `@BooleanValueMap`: Used to map a boolean value to a string value.

- `@ColorFilter`: Used to color a cell based on a value.

- `@ColumnTotal`: Used to generate a total column in the Excel file.

- `@CurrencyFormatter`: Used to format a number as currency.

- `@ExcelColumn`: Used to generate a column in the Excel file.

- `@ExcelCompatible`: Specifies that a specific class is compatible with Excel Generation.

- `@ExcelIntegerValueMap`: Used to map an integer value to a string value.

- `@ExcelLink`: Used to create a hyperlink in the Excel file.

- `@Formula`: Used to create a formula in the Excel file.

- `@Qualifier`: Enum used for cell color qualification.

- `@SheetFilter`: Used to use a field as a value for different sheets.

- `@StringMatcher`: Used to match a string value to a color.

- `@Style`: Used to style a cell.

## Basic Example

###Document reader/writer
```java
import org.apache.poi.ss.usermodel.Cell;
import za.co.silversurge.dataprocessing.ExcelFile;
import za.co.silversurge.gimcertificates.models.Document;

import java.io.EOFException;
import java.io.IOException;

public class DocumentExcel extends ExcelFile<Document> {
    public DocumentExcel(byte[] file) throws IOException {
        super(file);
    }

    public DocumentExcel(List<Document> data) {
        super(data, Document.class);
    }

    @Override
    protected void processCell(Document document, String s, Cell cell) throws EOFException {
        @Override
    protected void processCell(StudentPayment payment, String name, Cell cell) {
        if(cell == null)
            return;
        try {
            switch (name) {
                case "Document ID" -> {
                    document.setId((int)cell.getNumericCellValue());
                }
                case "Document Title" -> {
                    document.setDocumentTitle(cell.getStringCellValue());
                }
                case "Author" ->{
                    document.setAuthor(cell.getStringCellValue());
                }
                case "Creation Date" -> {
                    document.setCreationDate(cell.getDateCellValue());
                }
                case "Approved" -> {
                    document.setApproved(cell.getBooleanCellValue());
                }
            }
          }catch (Exception ex){
            throw new IllegalStateException("Column \"" + name + "\" is an invalid format in the template file");
        }
    }

    @Override
    public boolean isValid() {
        //TODO: add some logic to verify the data in the document.
        return true;
    }
}
```
### Class Implementation
```java
public class Document {

    @ExcelColumn(title = "Document ID", width = 20)
    private String documentId;

    @ExcelColumn(title = "Document Title", width = 30)
    private String documentTitle;

    @ExcelColumn(title = "Author", width = 25)
    private String author;

    @ExcelColumn(title = "Creation Date", format = "yyyy-MM-dd", width = 15)
    private Date creationDate;

    @ExcelColumn(title = "Is Approved", width = 15)
    @BooleanValueMap(trueValue = "Yes", falseValue = "No", matchColor = IndexedColors.GREEN, nonMatchColor = IndexedColors.RED)
    private boolean approved;

    // Other fields, getters, setters, and methods as needed

    public Document(String documentId, String documentTitle, String author, Date creationDate, boolean approved) {
        this.documentId = documentId;
        this.documentTitle = documentTitle;
        this.author = author;
        this.creationDate = creationDate;
        this.approved = approved;
    }

    // Getters and setters for the fields

    // Additional methods as needed
}
```
### Usage
#### Export

```java
// Example code to generate an Excel file
DocumentExcel excelFile = new DocumentExcel(/* List of document objects */);
final var excelBytes = excelFile.export()
```
#### Import
```java
DocumentExcel = excelFile = new DocumentExcel(/* Excel file byte array */);
List<String> errors = paymentExcel.getTableErrors();
//TODO: manage errors
final var documentList = paymentExcel.getTable();
```
