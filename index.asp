<%@ Language=VBScript %>
<!--#include file="conn.asp" -->
<%
Session("tested")=false
IF Session("error")=false THEN
	IF Application("teachername")="" THEN 
		Response.redirect("examsetting.asp")
	END IF
	dim came
	dim conn   
	dim connstr
	name=replace(trim(request("name")),"'","''")
	xuehao=replace(trim(request("xuehao")),"'","''")
	banji=replace(trim(request("banji")),"'","''")
	code=replace(trim(request("code")),"'","''")
	if(code=session("getcode"))then
		response.write("code correct")
		if ( name<>"" and xuehao <> "" and banji<>"" ) then
			set rs= server.createobject("adodb.recordset") 
			sql="select * from user"
			Set rs=conn.Execute(sql)
			rs.movefirst
			came=0
			Do while not rs.eof
				if rs("name")=name and rs("xuehao")=xuehao and rs("banji")=banji then
					came=1
				end if
				rs.movenext
			loop

			if came<> 1 then
				Response.write("err")
				Response.redirect "Error.asp"
			end if

			session("name")=name
			session("xuehao")=xuehao
			session("banji")=banji
			response.redirect "test.asp"
		else
			response.redirect "Error.asp"
		end if
	else
		if(code<>"") then
			response.redirect "Error.asp"
		end if
	end if
end if
Session("error")=false
%>
<meta charset="GBK">
<html>

<head>
	<meta NAME="GENERATOR" Content="Microsoft FrontPage 6.0">
	<title><%=Application("schoolname")%><%=Application("sitename")%>用户登记</title>
	<style>
		.jry_wb_button
		{
			background-color: #4CAF50;
			border: none;
			color: white;
			text-align: center;
			text-decoration: none;
			display: inline-block;
			-webkit-transition-duration: 0.4s;
			transition-duration: 0.4s;
			cursor: pointer;
		}
		.jry_wb_button_size_big,.jry_wb_button_size_big:focus
		{
			padding: 16px 32px;
			font-size: 23px;
			margin: 4px 2px;	
		}
		.jry_wb_button_size_big:hover
		{
			
		}
		.jry_wb_button_size_middle,.jry_wb_button_size_middle:focus
		{
			padding: 8px 16px;
			font-size: 23px;
			margin: 4px 2px;		
		}
		.jry_wb_button_size_middle:hover
		{

		}
		.jry_wb_button_size_small,.jry_wb_button_size_small:focus
		{
			padding: 2px 5px;
			font-size: 20px;
			margin: 4px 2px;		
		}
		.jry_wb_button_size_small:hover
		{
			
		}
		.jry_wb_color_normal{background-color:#40b3ce !important;color:#f2f2f2;}
		.jry_wb_color_normal_font{color:#40b3ce!important;}
		.jry_wb_color_normal:hover{background-color:#f2f2f2 !important;color:#40b3ce;}
		.jry_wb_color_normal_prevent:hover{background-color:#40b3ce !important;color:#f2f2f2;}
		.jry_wb_color_ok{background-color:#23dd07 !important;color:#f2f2f2;}
		.jry_wb_color_ok_font{color:#23dd07!important;}
		.jry_wb_color_ok:hover{background-color:#f2f2f2 !important;color:#23dd07;}
		.jry_wb_color_ok_prevent:hover{background-color:#23dd07 !important;color:#f2f2f2;}
		.jry_wb_color_warn{background-color:#ddcc00 !important;color:#f2f2f2;}
		.jry_wb_color_warn_font{color:#ddcc00!important;}
		.jry_wb_color_warn:hover{background-color:#f2f2f2 !important;color:#ddcc00;}
		.jry_wb_color_warn_prevent:hover{background-color:#ddcc00 !important;color:#f2f2f2;}
		.jry_wb_color_error{background-color:#ff0000 !important;color:#f2f2f2;}
		.jry_wb_color_error_font{color:#ff0000!important;}
		.jry_wb_color_error:hover{background-color:#f2f2f2 !important;color:#ff0000;}
		.jry_wb_color_error_prevent:hover{background-color:#ff0000 !important;color:#f2f2f2;}
	</style>
</head>

<body bgcolor="#999999" >

<table border="0" width="100%">
  <tr>
    <td width="100%">
      <p align="center" style="line-height: 150%">
		<font face="黑体" size="4" color="#33F">
		<span style="letter-spacing: 4pt"><%=Application("schoolname")%><%=Application("sitename")%></span></font></td>
  </tr>
</table>

<form action="index.asp" id="FORM1" method="post" name="FORM1">
	<table width="270" border="1" cellspacing="0" cellpadding="1" align="center" bordercolordark="#ecf5ff" bordercolorlight="#6699cc" height="1">
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">姓&nbsp;&nbsp;&nbsp;名：</td>       
			<td><input id="text1" name="name" style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">学&nbsp;&nbsp;&nbsp;号：</td>       
			<td><input id="password1" name="xuehao"  style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">班&nbsp;&nbsp;&nbsp;级：</td>       
			<td style="text-align:center;">
				<select size="1" name="banji" style="color:#33F;font-size:35px;">
					<option value="高一一班">高一一班</option>
					<option value="高一二班">高一二班</option>
					<option value="高一三班">高一三班</option>
					<option value="高一四班">高一四班</option>
					<option value="高一五班">高一五班</option>
					<option value="高一六班">高一六班</option>
					<option value="高一七班">高一七班</option>
					<option value="高一八班">高一八班</option>
					<option value="高一九班">高一九班</option>
					<option value="高一十班">高一十班</option>
					<option value="高一十一班">高一十一班</option>
					<option value="高一十二班">高一十二班</option>
					<option value="高一十三班">高一十三班</option>
					<option value="高一十四班">高一十四班</option>
					<option value="高二一班">高二一班</option>
					<option value="高二二班">高二二班</option>
					<option value="高二三班">高二三班</option>
					<option value="高二四班">高二四班</option>
					<option value="高二五班">高二五班</option>
					<option value="高二六班">高二六班</option>
					<option value="高二七班">高二七班</option>
					<option value="高二八班">高二八班</option>
					<option value="高二九班">高二九班</option>
					<option value="高二十班">高二十班</option>
					<option value="高二十一班">高二十一班</option>
					<option value="高二十二班">高二十二班</option>
					<option value="高二十三班">高二十三班</option>
					<option value="高二十四班">高二十四班</option>
				</select> 
			</td>
		</tr>				
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">人机身份验证：<img src="code.asp" /></td>    
			<td><input  id="p" name="code" style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr align="center">   
			<td height="5" width="128"><div align="center"><center><p><input type="submit" name="Submit1" value="进入" tabindex="5" class=" buttonface jry_wb_button jry_wb_button_size_big jry_wb_color_ok"></td>   
			<td height="21" width="128" align="center"><input type="reset" name="reset" value="重填" tabindex="6" class="buttonface jry_wb_button jry_wb_button_size_big jry_wb_color_warn"></td>   
		</tr>   
	</table>   
</form>   
<table border="1" cellspacing="0" cellpadding="1" align="center" bordercolordark="#ecf5ff" bordercolorlight="#6699cc" height="29">  
	<tr><td style="color:#33F;font-size:35px;text-align:center;"><a href="Order.asp">考试成绩名次查看</a></td></tr>   
	<tr><td style="color:#33F;font-size:35px;text-align:center;">若单击进入后页面被刷新，请检查您的验证码！</td></tr> 
</table>   
<!--#include file="info.asp" -->
</body>  
</html>  












