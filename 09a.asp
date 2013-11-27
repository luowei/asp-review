<html>
<head>
<title>第9题</title>
</head>
<script language="vbscript">
	function check_info()
		if form1.username.value="" then
			msgbox("用户名不能为空")
			check_info=0
		elseif form1.password.value="" then
			msgbox("密码不能为空")
			check_info=0
		elseif form1.email.value="" then
			msgbox("email不能为空")
			check_info=0
		else
		end if
	end function
</script>
<body>
<h1>第9题a</h1>
<h2>用户注册</h2>
	<form name="form1" onSubmit="check_info" action="" method="post">
	用户名：<input type="text" name="username"><br>
	密码：<input type="password" name="password"><br>
	性别：<input type="radio" name="sex" value="男" checked>男&nbsp;&nbsp;
	<input type="radio" name="sex" value="女">女<br>
	e-mail:<input type="text" name="email"><br>
	<input type="submit" value="提交">
	<input type="reset" value="重置">
	</form>
	<%
	if request.Form("username")<>"" And request.Form("password")<>"" then
		dim username,password,sex,email
		username=request.Form("username")
		password=request.Form("password")
		sex=request.Form("sex")
		email=request.Form("email")
		
		dim conn,strConn
		set conn=server.CreateObject("ADODB.Connection")
		'strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath("db.mdb")
		strConn="Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("db.mdb")
		conn.Open strConn
		
		Dim strSql
		strSql="insert into user(username,password,sex,email) values('"&username&"','"&password&"','"&sex&"','"&email&"')"
		conn.execute(strSql)
		response.Redirect "09b.asp"
	end if		
	%>
</body>
</html>