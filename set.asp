<%@ Language=VBScript %>

<%
Application("sitename")=request.form("sitename")
Application("teachername")=request.form("teachername")
Application("schoolname")=request.form("schoolname")
Application("q_type")=(request.form("q_type")) '''读取试题类型，1从头到尾读取，0随机
Application("q_num")=(request.form("q_num")) '''读取试题数量
Application("q_cost")=(request.form("q_cost")) '''每道题目占多少分
Application("q_show")=(request.form("q_show")) '''是否显示答案
Response.redirect("examsetting.asp?teachername="+Application("teachername"))
%>