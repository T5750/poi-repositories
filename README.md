# Apache POI

[![License](https://img.shields.io/badge/license-Apache-blue.svg)](https://github.com/T5750/poi/blob/master/LICENSE.txt)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/T5750/poi/pulls)
[![GitHub stars](https://img.shields.io/github/stars/T5750/poi.svg?style=social&label=Stars)](https://github.com/T5750/poi)
[![GitHub forks](https://img.shields.io/github/forks/T5750/poi.svg?style=social&label=Fork)](https://github.com/T5750/poi)

## Docs
- https://poix.readthedocs.io

## Getting Started
![](https://s0.wailian.download/2019/07/23/apache-poi-min-min.png)

Step 1: Download
```
git clone https://github.com/T5750/poi.git
cd poi
```

Step 2: Start Server
```
docker-compose up -d
# or
mvn clean spring-boot:run
```
[http://localhost:8080/poi](http://localhost:8080/poi)

### Test
- `TestReadExcel`, `TestReadExcelDemo`
- `TestExportExcel`, `TestExportExcel2007`, `TestWriteExcelDemo`
- `TestTemplate`, `TestExcelReplace`
- `CalendarDemo`
- `TestExcelFormulaDemo`, `TestExcelStylingDemo`, `TestAll`

### Runtime Environment
- [Java 8](https://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html)
- [Spring Framework 4.3.8.RELEASE](http://projects.spring.io/spring-framework)
- [Spring Boot 1.5.3.RELEASE](https://projects.spring.io/spring-boot)
- [Derby 10.13.1.1](https://db.apache.org/derby/)
- [Hibernate ORM 5.0.12.Final](http://hibernate.org/orm)
- [POI 3.17](https://poi.apache.org/download.html)
- [Bootstrap 4.3.1](https://github.com/twbs/bootstrap)
- [Docker 19.x](https://www.docker.com/)

### Classes
1. **HSSF, XSSF and XSSF classes**

Apache POI main classes usually start with either **HSSF**, **XSSF** or **SXSSF**.
- **HSSF** – is the POI Project’s pure Java implementation of the Excel ’97(-2007) file format. e.g. [HSSFWorkbook](https://poi.apache.org/apidocs/org/apache/poi/hssf/usermodel/HSSFWorkbook.html), [HSSFSheet](https://poi.apache.org/apidocs/org/apache/poi/hssf/usermodel/HSSFSheet.html).
- **XSSF** – is the POI Project’s pure Java implementation of the Excel 2007 OOXML (.xlsx) file format. e.g. [XSSFWorkbook](https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFWorkbook.html), [XSSFSheet](https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/XSSFSheet.html).
- **SXSSF** (since 3.8-beta3) – is an API-compatible streaming extension of XSSF to be used when very large spreadsheets have to be produced, and heap space is limited. e.g. [SXSSFWorkbook](https://poi.apache.org/apidocs/org/apache/poi/xssf/streaming/SXSSFWorkbook.html), [SXSSFSheet](https://poi.apache.org/apidocs/org/apache/poi/xssf/streaming/SXSSFSheet.html). SXSSF achieves its **low memory footprint by limiting access to the rows that are within a sliding window**, while XSSF gives access to all rows in the document.

2. **Row and Cell**

Apart from above classes, [Row](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Row.html) and [Cell](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/Cell.html) are used to interact with a particular row and a particular cell in excel sheet.

3. **Style Classes**

A wide range of classes like [CellStyle](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/CellStyle.html), [BuiltinFormats](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/BuiltinFormats.html), [ComparisonOperator](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/ComparisonOperator.html), [ConditionalFormattingRule](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/ConditionalFormattingRule.html), [FontFormatting](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/FontFormatting.html), [IndexedColors](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/IndexedColors.html), [PatternFormatting](https://poi.apache.org/apidocs/org/apache/poi/hssf/record/cf/PatternFormatting.html), [SheetConditionalFormatting](https://poi.apache.org/apidocs/org/apache/poi/ss/usermodel/SheetConditionalFormatting.html) etc. are used when you have to add formatting in a sheet, mostly based on some rules.

4. **FormulaEvaluator**

Another useful class **FormulaEvaluator** is used to evaluate the formula cells in excel sheet.

### Write an excel file
1. Create a workbook
1. Create a sheet in workbook
1. Create a row in sheet
1. Add cells in sheet
1. Repeat step 3 and 4 to write more data

### Read an excel file
1. Create workbook instance from excel sheet
1. Get to the desired sheet
1. Increment row number
1. iterate over all cells in a row
1. repeat step 3 and 4 until all data is read

## Getting Help
Having trouble with T5750's POI? We’d like to help!
- Ask a question on [CSDN](https://blog.csdn.net/evangel_z/article/details/7332535).
- Report bugs at https://github.com/T5750/poi/issues.

## Branch
View servlet branch at https://github.com/T5750/poi/tree/servlet.

## References
- [Java POI导出EXCEL经典实现 Java导出Excel弹出下载框](https://blog.csdn.net/evangel_z/article/details/7332535)
- [Java POI读取Office excel (2003,2007)及相关jar包](https://blog.csdn.net/evangel_z/article/details/7312050)
- [HSSF and XSSF Examples](http://poi.apache.org/spreadsheet/examples.html)
- [Apache POI – Read and Write Excel File in Java](https://howtodoinjava.com/library/readingwriting-excel-files-in-java-poi-tutorial/)

## License
This project is Open Source software released under the [Apache 2.0 license](http://www.apache.org/licenses/LICENSE-2.0.html).