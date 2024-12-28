# ecuacion-utils

## What is it?

`ecuacion-utils-poi` provides utilities for `apache-poi`.  
This library is dependent to `ecuacion-lib`.

## System Requirements

- JDK 21 or above.

## Documentation

- [javadoc](https://javadoc.ecuacion.jp/apidocs/ecuacion-util-poi/jp.ecuacion.util.poi/module-summary.html)

## Introduction

Check [Introduction](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page.  
dependency description is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-poi</artifactId>
    <version>4.0.0</version>
</dependency>
```

## features

We'll use the following table as an example. This table is in `Sheet1` sheet of `sample.xlsx`. The position of the top left cell is `A1`.

| name | age  | phone number   |
| ---- | ---- | ----           |
| John | 30   | (+01)123456789 |
| Ken  | 40   | (+81)987654321 |

### read values in excel cells as string

Following features read values of cells in excels and change into `String` datatype. Even if the value is defined as a number (like 12.3) in excel file, obtained values becomes `String`.  

#### read excel table values and put them to the list of strings

`SampleTableReader.java`

```java
public class SampleTableReader extends PoiStringFixedTableReader {

  public SampleTableReader() {
    
  }

  @Override
  protected String getSheetName() {
    return "Sheet1";
  }

  @Override
  protected String[] getHeaderLabels() {
    return new String[] {"name", "age", "phone number"};
  }

  public List<List<String>> read(String excelPath)
      throws EncryptedDocumentException, IOException, AppException {

     return getTableValues(excelPath);
  }
```
