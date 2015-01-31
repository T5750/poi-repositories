# Java POI 导入 导出 EXCEL
---
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
* 点击read POI，可以读取03和07版Excel
* 点击export POI，可以导出Excel
* 点击replace POI，读取模版`replaceTemplate.xls`，替换单元格内容后导出Excel
* 点击POI template，读取模版`template.xls`，插入新行后导出Excel

## Java Application

以下java类可直接运行使用

* read POI，对应`ReadExcel.java`
* export POI，对应`TestExportExcel.java`
* replace POI，对应`TestExcelReplace.java`
* POI template，对应`TestTemplate.java`

## Remarks

例子代码比较简单，仅供参考……

## Links

- [Java POI导出EXCEL经典实现 Java导出Excel弹出下载框](http://blog.csdn.net/evangel_z/article/details/7332535)

- [Java POI读取Office excel (2003,2007)及相关jar包](http://blog.csdn.net/evangel_z/article/details/7312050)

## Copyright

Copyright 2014-2015 evangel_z.
