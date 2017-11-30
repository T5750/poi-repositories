# Java POI 导入 导出 EXCEL

[![License](https://img.shields.io/badge/license-MIT-blue.svg)](https://github.com/T5750/poi/blob/master/LICENSE.md)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg)](https://github.com/T5750/poi/pulls)
[![GitHub stars](https://img.shields.io/github/stars/T5750/poi.svg?style=social&label=Stars)](https://github.com/T5750/poi)
[![GitHub forks](https://img.shields.io/github/forks/T5750/poi.svg?style=social&label=Fork)](https://github.com/T5750/poi)

## Runtime Environment

* Java version: 1.6.0_33
* Eclipse Java EE IDE for Web Developers Version: Luna Service Release 1 (4.4.1)
* Text file encoding: UTF-8
* Server version: Apache Tomcat/6.0.41
* Git version: 1.7.9.5

## Java Web

* 将源代码导入eclipse中
* 部署项目，启动tomcat服务器
* web页面具体路径：http://localhost:8080/poi
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

例子代码比较简单，仅供参考……

## Links

- [Java POI导出EXCEL经典实现 Java导出Excel弹出下载框](http://blog.csdn.net/evangel_z/article/details/7332535)
- [Java POI读取Office excel (2003,2007)及相关jar包](http://blog.csdn.net/evangel_z/article/details/7312050)
- [HSSF and XSSF Examples](http://poi.apache.org/spreadsheet/examples.html)

## Copyright

Copyright 2014-2017 evangel_z.
