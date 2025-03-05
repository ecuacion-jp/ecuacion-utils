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
- `jakarta.annotation:jakarta.annotation-api`
- `org.slf4j:slf4j-api`

(modules depending on `ecuacion-lib-core`)
- `jakarta.mail:jakarta.mail-api` (If you want to use the mail related utility: `jp.ecuacion.lib.core.util.MailUtil`)
- `org.hibernate.validator:hibernate-validator`
- `org.glassfish:jakarta.el`
- (any logging libraries. `ch.qos.logback:logback-classic` is reccomended.)

## Documentation

- [javadoc](https://javadoc.ecuacion.jp/apidocs/ecuacion-util-poi/jp.ecuacion.util.poi/module-summary.html)

## Sample Code

- [ecuacion-util-poi-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-poi-sample)

## Introduction

Check [Introduction](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
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
