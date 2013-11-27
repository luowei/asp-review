<html>
<head>
	<title>第5题</title>
</head>
<body>
	<%
		response.Write "当前cookies中的姓名是：" & request.Cookies("user")("username")
		if Request.Form("username")<>"" and Request.Form("age")<>"" then
			Response.Cookies("user")("username")=Request.Form("username")
			response.cookies("user")("age")=Request.form("age")
			response.cookies("user").expires=DateAdd("d",Request.Form("validity"),Date())
		end if
	%>
	<form name="form1" action="" method="post">
	姓名：<input type="text" name="username"><br>
	年龄：<input type="text" name="age"><br>
	请选择cookie有效期：
	<input type="radio" name="validity" value="7">一周 &nbsp;&nbsp;
	<input type="radio" name="validity" value="30">一月 &nbsp;&nbsp;
	<input type="radio" name="validity" value="365">一年 &nbsp;&nbsp;
	<input type="submit" value="确定">
	</form>
</body>
</html>