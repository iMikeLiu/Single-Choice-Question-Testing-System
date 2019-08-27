<meta charset="GBK">
<%@ Language=VBScript %>
<%
IF Session("tested")=true THEN
	Response.redirect "index.asp"
END IF
%>
<!--#include file="conn.asp" -->
<%
pick=1
lowerb=1
upperb=24
randomize()
pick=Int((upperb-lowerb+1)*Rnd+lowerb)
session("pick_m")=pick
%>
<%
if Application("q_type")=1 then
sql="select top "&Application("q_num")&" * from test"
else
sql="select * from test where ( num >=" & pick & " and num <=" & (pick+ Application("q_num"))-1 & ")"
end if
Set rs = conn.Execute(sql)
%>
<html>
<head>
<style type="text/css">
<!--
.unnamed1 {  font-size: 10pt; text-decoration: none}
-->
</style>
<title><%=schoolname%><%=sitename%></title>
</head>
<body bgcolor="#FFFFFF">
<table border="0" width="100%">
  <tr>
    <td width="100%"><p align="center" style="line-height: 150%">
		<font face="黑体" size="4" color="#000080">
		<span style="letter-spacing: 4pt"><%=schoolname%><%=sitename%></span></font></td>
  </tr>
</table>
<form name="forms">
  <span class="unnamed1"><table border="1" width="100%"
  height="1" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
    <tr>
      <td width="25%" height="8" align="center"><p><b><font color="#008000" size="3">姓&nbsp;&nbsp; 名</font></b>     
      </td>     
      <td width="25%" height="8" align="center"><p><b><font color="#008000" size="3">学&nbsp;&nbsp; 号</font></b>     
      </td>     
      <td width="25%" height="8" align="center"><p><b><font color="#008000" size="3">班&nbsp;&nbsp; 级</font></b>     
      </td>     
      <td width="25%" height="8" align="center"><p><span     
      class="unnamed1"><b><font color="#008000" size="3">考试时间</font></b></span>     
      </td>     
    </tr>     
    <tr align="center">     
      <td width="25%" height="1"><font size="2"><%=session("name")%></font> </td>     
      <td width="25%" height="1"><font size="2"><%=session("xuehao")%></font> </td>     
      <td width="25%" height="1"><font size="2"><%=session("banji")%></font> </td>     
      <td width="25%" height="1"><center><p><span class="unnamed1"><input     
      type="text" name="input1" size="9"> <script language="javascript"><!--     
     
var sec=0;var min=0;var hou=0;flag=0;idt=window.setTimeout("update();",1000);function update(){sec++;if(sec==60){sec=0;min+=1;}if(min==60){min=0;hou+=1;}if((min>0)&&(flag==0)){flag=1;}     
     
document.forms.input1.value=hou+"时"+min+"分"+sec+"秒";idt=window.setTimeout("update();",1000);};     
//-->     
</script> </span></td>     
    </tr>     
  </table>     
  </center></span>     
     
</form>     
<%i=1     
rs.movefirst     
do while not rs.eof     
%>     
<form action="result.asp" id="FORM1" method="post" name="FORM1">     
<p><span class="unnamed1"></span></p>
<table align="center" border="1" cellPadding="0"    
  cellSpacing="0" width="100%" bordercolor="#C0C0C0" bgcolor="#ABF2F5" height="37" bordercolorlight="#008000" bordercolordark="#FFFFFF"> 
    <tr bgcolor="#66CCFF">    
      <td width="51%" class="unnamed1" bgcolor="#ffffff" height="15" colspan="4">
       <b><%="("&i&")"%> <left> <%=rs("question")%></b></td>    
    </tr>    
      <tr bgcolor="#99CCFF">    
      <td width="24%" class="unnamed1" bgcolor="#ABF2F5" height="1"><p align="left"><input name="ans<%=i%>" type="radio" value="A"> <%=rs("A")%></p> </td>     
      <td width="25%" class="unnamed1" bgcolor="#ABF2F5" height="1"><input name="ans<%=i%>" type="radio"     
      value="B"> <%=rs("B")%></td>     
      <td width="25%" class="unnamed1" bgcolor="#ABF2F5" height="1"><input name="ans<%=i%>" type="radio"     
      value="C"> <%=rs("C")%></td>     
      <td width="25%" class="unnamed1" bgcolor="#ABF2F5" height="1"><input name="ans<%=i%>" type="radio"     
      value="D"> <%=rs("D")%></td>     
    </tr>     
  </table>     
  <span class="unnamed1"><%     
i=i+1     
rs.movenext     
loop     
%>     
</span><center>
  <p><span class="unnamed1"></span></p>     
  </center><center><p></p></center>
    <center><p><span class="unnamed1"> 
     <input id="submit1" name="submit1" type="submit" value="提交">
     <input id="reset1" name="reset1" type="reset" value="重填"> </span></p>     
    </center>     
</form> 
</body>     
</html>