<html>
<head>
	<title>��7��</title>
</head>
<body>
	<form name="form1" method="post" action="">
	������<input type="text" name="username"><br>
	�������ݣ�<br><textarea name="say" rows="10" cols="40">����������...</textarea><br>
	<input type="submit" value="����"><br>
	</form>
	<%
	if trim(request.Form("username"))<>"" and trim(request.Form("say"))<>"" and trim(request.Form("say"))<>"����������..." then
		Dim say
		say="IPΪ��" & request.ServerVariables("REMOTE_ADDR") & "<br>" & request.Form("username") & "<br>" & Now() & "˵��" & request.Form("say") & "<br><br>"
		application.Lock()
		application("strChat")=application("strChat") & say
		application.UnLock()
		response.Write application("strchat")
	end if
	%>
	
</body>
</html>