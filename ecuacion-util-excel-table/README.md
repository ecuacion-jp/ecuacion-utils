# ecuacion-util-excel-table

## What is it?

`ecuacion-util-excel-table` lets you read and write structured table data in Excel (`.xlsx`) files without writing raw Apache POI boilerplate. Describe the table by its sheet name and header labels — the library locates the data rows, maps each row to a Java bean, and handles all cell-type conversions.

## Usage Example

```java
List<MemberBean> members = new StringOneLineHeaderExcelTableToBeanReader<>(
    MemberBean.class,
    "Member",                                                    // sheet name
    new String[] {"ID", "name", "date of birth", "age"})        // header labels to match
        .tableStartRowNumber(3)
        .tableStartColumnNumber(2)
        .readToBean("members.xlsx");
```

That's all. Name the sheet, list the expected headers, and get a typed list back.

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
- `org.slf4j:slf4j-api`
- (if you use log4j2, add `org.apache.logging.log4j:log4j-slf4j-impl` and `org.apache.logging.log4j:log4j-core`,
   or else `org.apache.logging.log4j.log4j-to-slf4j` (To use any slf4j-compatible logging modules) and any `org.slf4j:slf4j-api` compatible logging libraries. `ch.qos.logback:logback-classic` is recommended.)

(modules depending on `ecuacion-lib-core`)
- `jakarta.mail:jakarta.mail-api` (If you want to use the mail related utility: `jp.ecuacion.lib.core.util.MailUtil`)

Since the dependency libraries are a little complicated, we recommend to refer `pom.xml` in `ecuacion-util-excel-table-sample`. 

## Documentation

- [javadoc](https://javadoc.io/doc/jp.ecuacion.util/ecuacion-util-excel-table/latest/jp.ecuacion.util.excel/module-summary.html)

## Sample Code

- [ecuacion-util-excel-table-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-excel-table-sample)

## Installation

Check [Installation](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
The description of dependent `ecuacion` modules is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-excel-table</artifactId>
    <!-- Put the latest release version -->
    <version>x.x.x</version>
</dependency>

<!-- ecuacion-lib -->
<dependency>
    <groupId>jp.ecuacion.lib</groupId>
    <artifactId>ecuacion-lib-validation</artifactId>
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

#### Read table values as strings

```java
List<List<String>> rows = new StringOneLineHeaderExcelTableReader(
    "Sheet1", new String[]{"name", "age", "phone number"})
    .read("sample.xlsx");
```

Each inner list contains the values of one data row in the header order.

#### Read table values into Java beans

```java
List<PersonBean> people = new StringOneLineHeaderExcelTableToBeanReader<>(PersonBean.class,
    "Sheet1", new String[]{"name", "age", "phone number"})
    .readToBean("sample.xlsx");
```

For more examples — free-format tables, cell-level access, writing — see [Sample Code](#sample-code) above.
