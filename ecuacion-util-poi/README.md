# ecuacion-util-poi

## What is it?

`ecuacion-util-poi` provides utilities for `apache POI`.  

## System Requirements

- JDK 21 or above.

## Dependent Ecuacion Libraries

### Automatically Loaded Libraries

(none)

### Manual Load Needed Libraries

- `ecuacion-lib-core`

## Dependent External Libraries

### Automatically Loaded Libraries

- `org.apache.poi:poi`
- `org.apache.poi:poi-ooxml`

### Manual Load Needed Libraries

- `jakarta.validation:jakarta.validation-api`
- (any `jakarta.validation:jakarta.validation-api` compatible Jakarta Validation libraries. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- `jakarta.annotation:jakarta.annotation-api`
- `org.slf4j:slf4j-api`
- (if you use log4j2, add `org.apache.logging.log4j:log4j-slf4j-impl` and `org.apache.logging.log4j:log4j-core`,
   or else `org.apache.logging.log4j.log4j-to-slf4j` (To use any slf4j-compatible logging modules) and any `org.slf4j:slf4j-api` compatible logging libraries. `ch.qos.logback:logback-classic` is recommended.)

(modules depending on `ecuacion-lib-core`)
- `jakarta.mail:jakarta.mail-api` (If you want to use the mail related utility: `jp.ecuacion.lib.core.util.MailUtil`)

Since the dependency libraries are a little complicated, we recommend to refer `pom.xml` in `ecuacion-util-poi-sample`. 

## Documentation

- [javadoc](https://javadoc.ecuacion.jp/apidocs/ecuacion-util-poi/jp.ecuacion.util.poi/module-summary.html)

## Sample Code

- [ecuacion-util-poi-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-poi-sample)

## Installation

Check [Installation](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
The description of dependent `ecuacion` modules is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-poi</artifactId>
    <!-- Put the latest release version -->
    <version>x.x.x</version>
</dependency>

<!-- ecuacion-lib -->
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-lib-core</artifactId>
    <!-- Put the latest release version -->
    <version>x.x.x</version>
</dependency>
```

## Features

We'll use the following table as an example. Let's say this table is in `Sheet1` sheet of `sample.xlsx`, the position of the top left cell of the table is `A1`.

| name | age  | phone number   |
| ---- | ---- | ----           |
| John | 30   | (+01)123456789 |
| Ken  | 40   | (+81)987654321 |

### Read Values In Excel Cells As String

The following features read values of cells in the excel file and change into `String` datatype. Even if the value is defined as a number (like 12.3) in excel file, obtained values becomes `String`.  

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
