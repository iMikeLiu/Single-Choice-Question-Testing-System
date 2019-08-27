<!--#include file="conn.asp" -->
<%
'set rs= server.createobject("adodb.recordset") 
sql="delete from user where name = '"&request.QueryString("word")+"'"
Set rs=conn.Execute(sql)
%>