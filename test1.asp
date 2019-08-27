<html>
<%

	Randomize
	Dim Str 

	Const numCode = "0123456789"
	Const optCode = "£«£­¡Á¡Â"
	For i = 1 to Int(Rnd * 10) 
		Str(i)=numCode(i)
	Next
	Response.write(Str)

%>
</html>