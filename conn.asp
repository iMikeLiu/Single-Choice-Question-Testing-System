<%
connstr="DBQ="+server.mappath("exercise.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
%>