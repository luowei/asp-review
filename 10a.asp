<html>
<head>
<title>第10题</title>
</head>
<body>
<h1>第10题a</h1>
<script language="vbscript">
	function check_info()
		if form1.username.value="" then
			msgbox "用户名不能为空"
			check_info=0
		elseif form1.password.value="" then
			msgbox "密码不能为空"
			check_info=0
		elseif form1.email.value="" then
			msgbox "email地址不能为空"
			check_info=0
		else
			chck_info=1
		end if
	end function
</script>
<form name="form1" onSubmit="check_info" method="post" action="">
用户名：<input type="text" name="username"><br>
密码：<input type="password" name="password"><br>
性别：
<input type="radio" name="sex" value="男" checked>男&nbsp;&nbsp;
<input type="radio" name="sex" value="女">女<br>
email:<input type="text" name="email"><br>
<input type="submit" value="提交">&nbsp;&nbsp;
<input type="reset" value="重置">
</form>
<%
if not isempty(request.Form("username")) then
	dim  username,password,sex,email
	username=trim(request.Form("username"))
	password=trim(request.Form("password"))
	sex=request.Form("sex")
	email=trim(request.Form("email"))
	
	dim conn,strConn
	set conn=server.CreateObject("ADODB.Connection")
	strConn="Driver={Microsoft Access Driver (*.mdb)};Dbq=" &server.MapPath("db.mdb")
	conn.open strConn
	
	dim strSql
	strSql="insert into user(username,password,sex,email) values('"&username&"','"&password&"','"&sex&"','"&email&"')"
	conn.execute(strSql)
	response.Write "已向user表中了写入了"
	response.Write(username)
end if
%>

</body>
</html>