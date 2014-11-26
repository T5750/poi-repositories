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
function readExcel(){
	var postForm = document.getElementById("excelForm");
	postForm.action = "<%=basePath%>read";
	postForm.submit();
}
function exportExcel(){
	var postForm = document.getElementById("excelForm");
	postForm.action = "<%=basePath%>export";
	postForm.submit();
}
</script>
  </head>
  <body>
  	<form id="excelForm" method="post">
	    <a href="javascript:void(0)" onclick="readExcel()">read POI</a><br>
	    <a href="javascript:void(0)" onclick="exportExcel()">export POI</a><br>
    </form>
  </body>
</html>
