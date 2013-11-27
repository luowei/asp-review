<html>
<head>
<title>第9题</title>
</head>
<body>
<h1>第9题b</h1>
<h2>登录验证</h2>
	<form name="form1" action="" method="post">
	用户名：<input type="text" name="username"><br>
	密码：<input type="password" name="password"><br>
	<input type="submit" value="登录">
	<input type="reset" value="重置">
	</form><br>
<%
if IsEmpty(session("username")) then
	session("username")=""
elseif session("username")="" then
	dim username,password
	username=trim(request.Form("username"))
	password=trim(request.Form("password"))
	
	dim sql,conn,rs
	sql="select * from user where username='"&username&"'"
	set conn=server.CreateObject("ADODB.Connection")
	'strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath("db.mdb")
	strConn="Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("db.mdb")
	conn.open strConn
	Set rs=conn.execute(sql)

	if rs.eof then
		response.Write "<br>登录失败</body></html>"
		response.End()
	elseif rs("password")<>password then
		response.Write "<br>登录失败</body></html>"
		response.End()
	else
		session("username")=username
		response.Write "<br>登录成功</body></html>"
	end if
else
end if
%>
</body>
</html>