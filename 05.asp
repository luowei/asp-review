<html>
<head>
	<title>��5��</title>
</head>
<body>
	<%
		response.Write "��ǰcookies�е������ǣ�" & request.Cookies("user")("username")
		if Request.Form("username")<>"" and Request.Form("age")<>"" then
			Response.Cookies("user")("username")=Request.Form("username")
			response.cookies("user")("age")=Request.form("age")
			response.cookies("user").expires=DateAdd("d",Request.Form("validity"),Date())
		end if
	%>
	<form name="form1" action="" method="post">
	������<input type="text" name="username"><br>
	���䣺<input type="text" name="age"><br>
	��ѡ��cookie��Ч�ڣ�
	<input type="radio" name="validity" value="7">һ�� &nbsp;&nbsp;
	<input type="radio" name="validity" value="30">һ�� &nbsp;&nbsp;
	<input type="radio" name="validity" value="365">һ�� &nbsp;&nbsp;
	<input type="submit" value="ȷ��">
	</form>
</body>
</html>