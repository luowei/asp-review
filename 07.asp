<html>
<head>
	<title>第7题</title>
</head>
<body>
	<form name="form1" method="post" action="">
	姓名：<input type="text" name="username"><br>
	发言内容：<br><textarea name="say" rows="10" cols="40">请输入内容...</textarea><br>
	<input type="submit" value="发送"><br>
	</form>
	<%
	if trim(request.Form("username"))<>"" and trim(request.Form("say"))<>"" and trim(request.Form("say"))<>"请输入内容..." then
		Dim say
		say="IP为：" & request.ServerVariables("REMOTE_ADDR") & "<br>" & request.Form("username") & "<br>" & Now() & "说：" & request.Form("say") & "<br><br>"
		application.Lock()
		application("strChat")=application("strChat") & say
		application.UnLock()
		response.Write application("strchat")
	end if
	%>
	
</body>
</html>