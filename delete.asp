

<%@ Language=VBScript %>
<!--#include file="conn.asp" -->
<html>
<head>
<title>
管理员删除用户模块
</title>
<meta http-equiv="Content-Type" content="text/html; charset=GBK">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="Microsoft Theme" content="none"></head>
<body>
<%
set rs= server.createobject("adodb.recordset") 
sql="select * from user"
Set rs=conn.Execute(sql)
if(rs.eof=false)then
rs.movefirst
Do while not rs.eof
%>
<p><a href="del.asp?word=<%=rs("name")%>"><%=rs("name")%></a>
</p>
<%
rs.movenext
loop
else
response.write "无用户！"
END IF
%>

</body>
</html>