<%@page import="com.candidjava.ExcelDataReterive"%>
<%@page import="java.util.List;"%>

<%@ page language="java" contentType="text/html; charset=ISO-8859-1"
    pageEncoding="ISO-8859-1"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1">
<title>Display</title>
<style>
table#nat{
	width: 50%;
	background-color: #c48ec5;
}
</style>
</head>
<body>  

<table id ="nat">

<tr>
	<td>Date</td>
	<td>Name</td>
	<td>Contact</td>
	<td>Course</td>
	<td>Full Fee</td>
	<td>Paid</td>
	<td>Balance</td>
	
</tr>

<% String name =  request.getParameter("fullname");
	String number = request.getParameter("number");
	List<String> data;
	if(name!=null){
		data = ExcelDataReterive.getDatafromExcel("Name", name);
	}else{
		data = ExcelDataReterive.getDatafromExcel("Number", number);
	}
	
for(int i=0; i<data.size(); i++){	
	String source  = data.get(i);
	String[] sp = source.split("&");
	
	%>
<tr>
   <td>
      <%= sp[0]%>
   </td>
    <td>
      <%= sp[1]%>
   </td>
    <td>
      <%= sp[2]%>
   </td>
    <td>
      <%= sp[3]%>
   </td>
    <td>
      <%= sp[4]%>
   </td>
    <td>
      <%= sp[5]%>
   </td>
      <td>
      <%= sp[6]%>
   </td>
   </tr>
<% } %>



</table>
</body>
</html>