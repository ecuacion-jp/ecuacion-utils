# ecuacion-util-excel-table-sample

## What is it?

Sample code for [`ecuacion-util-excel-table`](../ecuacion-util-excel-table).

The `pom.xml` in this project also serves as a reference for dependency setup.

## Sample Classes

| Package | Class | Description |
| --- | --- | --- |
| `readstringfromcell` | `ReadStringsFromHeaderFormatExcelTableToBeanSample` | Reads rows from a header-format table into a list of beans |
| `copytable` | `CopyHeaderFormatExcelTableSample` | Reads a header-format table and writes it to another sheet |
| `copytable` | `IterativeCopyHeaderFormatExcelTableSample` | Same as above, but writes row by row (memory-efficient for large tables) |
| `copytable` | `CopyFreeFormatExcelTableSample` | Reads a free-format table and writes it to another sheet |
