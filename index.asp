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
	<title><%=Application("schoolname")%><%=Application("sitename")%>�û��Ǽ�</title>
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
		<font face="����" size="4" color="#33F">
		<span style="letter-spacing: 4pt"><%=Application("schoolname")%><%=Application("sitename")%></span></font></td>
  </tr>
</table>

<form action="index.asp" id="FORM1" method="post" name="FORM1">
	<table width="270" border="1" cellspacing="0" cellpadding="1" align="center" bordercolordark="#ecf5ff" bordercolorlight="#6699cc" height="1">
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">��&nbsp;&nbsp;&nbsp;����</td>       
			<td><input id="text1" name="name" style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">ѧ&nbsp;&nbsp;&nbsp;�ţ�</td>       
			<td><input id="password1" name="xuehao"  style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">��&nbsp;&nbsp;&nbsp;����</td>       
			<td style="text-align:center;">
				<select size="1" name="banji" style="color:#33F;font-size:35px;">
					<option value="��һһ��">��һһ��</option>
					<option value="��һ����">��һ����</option>
					<option value="��һ����">��һ����</option>
					<option value="��һ�İ�">��һ�İ�</option>
					<option value="��һ���">��һ���</option>
					<option value="��һ����">��һ����</option>
					<option value="��һ�߰�">��һ�߰�</option>
					<option value="��һ�˰�">��һ�˰�</option>
					<option value="��һ�Ű�">��һ�Ű�</option>
					<option value="��һʮ��">��һʮ��</option>
					<option value="��һʮһ��">��һʮһ��</option>
					<option value="��һʮ����">��һʮ����</option>
					<option value="��һʮ����">��һʮ����</option>
					<option value="��һʮ�İ�">��һʮ�İ�</option>
					<option value="�߶�һ��">�߶�һ��</option>
					<option value="�߶�����">�߶�����</option>
					<option value="�߶�����">�߶�����</option>
					<option value="�߶��İ�">�߶��İ�</option>
					<option value="�߶����">�߶����</option>
					<option value="�߶�����">�߶�����</option>
					<option value="�߶��߰�">�߶��߰�</option>
					<option value="�߶��˰�">�߶��˰�</option>
					<option value="�߶��Ű�">�߶��Ű�</option>
					<option value="�߶�ʮ��">�߶�ʮ��</option>
					<option value="�߶�ʮһ��">�߶�ʮһ��</option>
					<option value="�߶�ʮ����">�߶�ʮ����</option>
					<option value="�߶�ʮ����">�߶�ʮ����</option>
					<option value="�߶�ʮ�İ�">�߶�ʮ�İ�</option>
				</select> 
			</td>
		</tr>				
		<tr>
			<td style="color:#33F;font-size:35px;text-align:center;">�˻������֤��<img src="code.asp" /></td>    
			<td><input  id="p" name="code" style="color:#33F;font-size:35px;text-align:center;" size="20" tabindex="1"/></td>       
		</tr>
		<tr align="center">   
			<td height="5" width="128"><div align="center"><center><p><input type="submit" name="Submit1" value="����" tabindex="5" class=" buttonface jry_wb_button jry_wb_button_size_big jry_wb_color_ok"></td>   
			<td height="21" width="128" align="center"><input type="reset" name="reset" value="����" tabindex="6" class="buttonface jry_wb_button jry_wb_button_size_big jry_wb_color_warn"></td>   
		</tr>   
	</table>   
</form>   
<table border="1" cellspacing="0" cellpadding="1" align="center" bordercolordark="#ecf5ff" bordercolorlight="#6699cc" height="29">  
	<tr><td style="color:#33F;font-size:35px;text-align:center;"><a href="Order.asp">���Գɼ����β鿴</a></td></tr>   
	<tr><td style="color:#33F;font-size:35px;text-align:center;">�����������ҳ�汻ˢ�£�����������֤�룡</td></tr> 
</table>   
<!--#include file="info.asp" -->
</body>  
</html>  












