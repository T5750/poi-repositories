# Java POI 导入 导出 EXCEL

[![License](https://img.shields.io/badge/license-Apache-blue.svg)](https://github.com/T5750/poi/blob/master/LICENSE.txt)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/T5750/poi/pulls)
[![GitHub stars](https://img.shields.io/github/stars/T5750/poi.svg?style=social&label=Stars)](https://github.com/T5750/poi)
[![GitHub forks](https://img.shields.io/github/forks/T5750/poi.svg?style=social&label=Fork)](https://github.com/T5750/poi)

## Runtime Environment
- [Java 6](http://www.oracle.com/technetwork/java/javase/downloads/jdk6downloads-1902814.html)
- [Tomcat 6](http://tomcat.apache.org/)
- [Eclipse IDE for Java EE Developers](http://www.eclipse.org/downloads/eclipse-packages/)
- [IntelliJ IDEA](https://www.jetbrains.com/idea/download/)
- Text file encoding: UTF-8

## Java Web
* 将源代码导入Eclipse或IDEA中
* 部署项目，启动Tomcat服务器
* Web页面具体路径：[http://localhost:8080/poi](http://localhost:8080/poi)
* 点击Read excel 2003 or 2007，可以读取2003或2007版Excel
* 点击Export excel 2003，可以导出Excel 2003版
* 点击Export excel 2007，可以导出Excel 2007版
* 点击Export excel by templates,replace a cell，读取模版`replaceTemplate.xls`，替换单元格内容后导出Excel
* 点击Export excel by templates,insert a table，读取模版`template.xls`，插入表格后导出Excel

## Java Application
以下java类可直接运行使用

* Read excel 2003 or 2007，对应`ReadExcel.java`
* Export excel 2003，对应`TestExportExcel.java`
* Export excel 2007，对应`TestExportExcel2007.java`
* Export excel by templates,replace a cell，对应`TestExcelReplace.java`
* Export excel by templates,insert a table，对应`TestTemplate.java`

## Remarks
- 例子代码比较简单，仅供参考……
- 兼容Java 7, Java 8

## Links
- [Java POI导出EXCEL经典实现 Java导出Excel弹出下载框](http://blog.csdn.net/evangel_z/article/details/7332535)
- [Java POI读取Office excel (2003,2007)及相关jar包](http://blog.csdn.net/evangel_z/article/details/7312050)
- [HSSF and XSSF Examples](http://poi.apache.org/spreadsheet/examples.html)

## License
This project is Open Source software released under the [Apache 2.0 license](http://www.apache.org/licenses/LICENSE-2.0.html).
