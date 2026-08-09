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
- `jakarta.validation:jakarta.validation-api`
- `org.slf4j:slf4j-api`
- `org.apache.commons:commons-lang3`

### Manual Load Needed Libraries

- (If you use validation features, any `jakarta.validation:jakarta.validation-api` compatible Jakarta Validation implementation is needed. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- (If you use `MailUtil` (in `ecuacion-lib-core`), `jakarta.mail:jakarta.mail-api` and any compatible Jakarta Mail implementation are needed. `org.eclipse.angus:angus-mail` is recommended.)
- (any `org.slf4j:slf4j-api` compatible logging implementation. `ch.qos.logback:logback-classic` is recommended.)

## Documentation

- [ecuacion-references-utils](https://references.ecuacion.jp/ecuacion-references-utils/public/showMarkdown/page?id=home) — Official reference documentation
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
