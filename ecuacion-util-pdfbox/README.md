# ecuacion-util-poi

## What is it?

`ecuacion-util-pdfbox` provides utilities for `apache PDFBox`.  

## System Requirements

- JDK 21 or above.

## Dependent Ecuacion Libraries

### Automatically Loaded Libraries

(none)

### Manual Load Needed Libraries

- `ecuacion-lib-core`

## Dependent External Libraries

### Automatically Loaded Libraries

- `org.apache.pdfbox`
- `org.apache.poi:poi`
- `org.apache.poi:poi-ooxml`

### Manual Load Needed Libraries

- `jakarta.validation:jakarta.validation-api`
- (any `jakarta.validation:jakarta.validation-api` compatible Bean Validation libraries. `org.hibernate.validator:hibernate-validator` and `org.glassfish:jakarta.el` are recommended.)
- `jakarta.annotation:jakarta.annotation-api`
- `org.slf4j:slf4j-api`
- (if you use log4j2, add `org.apache.logging.log4j:log4j-slf4j-impl` and `org.apache.logging.log4j:log4j-core`,
   or else `org.apache.logging.log4j.log4j-to-slf4j` (To use any slf4j-compatible logging modules) and any `org.slf4j:slf4j-api` compatible logging libraries. `ch.qos.logback:logback-classic` is recommended.
(modules depending on `ecuacion-lib-core`)
- `jakarta.mail:jakarta.mail-api` (If you want to use the mail related utility: `jp.ecuacion.lib.core.util.MailUtil`)

Since the dependency libraries are a little complicated, we recommend to refer `pom.xml` in `ecuacion-util-pdfbox-sample`. 

## Documentation

- [javadoc](https://javadoc.ecuacion.jp/apidocs/ecuacion-util-pdfbox/)

## Sample Code

- [ecuacion-util-pdfbox-sample](https://github.com/ecuacion-jp/ecuacion-utils/tree/main/ecuacion-util-pdfbox-sample)

## Introduction

Check [Introduction](https://github.com/ecuacion-jp/ecuacion-lib) part of `README` page in `ecuacion-lib`.  
The description of dependent `ecuacion` modules is as follows.

```xml
<dependency>
    <groupId>jp.ecuacion.util</groupId>
    <artifactId>ecuacion-util-pdfbox</artifactId>
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

### Create PDF from Excel

#### Excel Values used to create PDF files

- File > Page Setup > Page tab > orientation
- File > Page Setup > Page tab > margins
