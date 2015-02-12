<%@ page language="java" import="java.util.*" pageEncoding="UTF-8"%>
<%
String path = request.getContextPath();
String basePath = request.getScheme()+"://"+request.getServerName()+":"+request.getServerPort()+path+"/";
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
  <head>
    <base href="<%=basePath%>">
    <title>POI</title>
	<meta http-equiv="pragma" content="no-cache">
	<meta http-equiv="cache-control" content="no-cache">
	<meta http-equiv="expires" content="0">    
	<meta http-equiv="keywords" content="keyword1,keyword2,keyword3">
	<meta http-equiv="description" content="This is my page">
	<!--
	<link rel="stylesheet" type="text/css" href="styles.css">
	-->
<script type="text/javascript">
function formSubmit(mapping){
	var postForm = document.getElementById("excelForm");
	postForm.action = "<%=basePath%>" + mapping;
	postForm.submit();
}
</script>
  </head>
  <body>
  	<h2>POI</h2><hr><br>
  	<form id="excelForm" method="post">
	    <a href="javascript:void(0)" onclick="formSubmit('read')">Read excel 2003 or 2007</a><br><br>
	    <a href="javascript:void(0)" onclick="formSubmit('export')">Export excel 2003</a><br><br>
	    <a href="javascript:void(0)" onclick="formSubmit('export2007')">Export excel 2007</a><br><br>
	    <a href="javascript:void(0)" onclick="formSubmit('replace')">Export excel by templates,replace a cell</a><br><br>
	    <a href="javascript:void(0)" onclick="formSubmit('template')">Export excel by templates,insert a table</a><br><br>
    </form>
  </body>
</html>
