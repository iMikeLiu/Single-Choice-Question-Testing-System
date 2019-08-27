<html>
<p style="color:blue">您的学号和姓名、班级不匹配！3秒钟后会重定向至主页！</p>
<%
Session("error")=true
%>
<%Response.redirect("index.asp")%>
</html>