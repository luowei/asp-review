<html>
<head>
<title>第11题</title>
</head>
<body>
<center>
<h1>第11题b</h1>
<table align="center" border="1">
<tr><th>用户名</th><th>密码</th><th>性别</th><th>email</th></tr>
<%
	'if isempty(request.Form("username")) then	
	dim conn,username
	username=request.QueryString("username")
	response.Write "删除"&username&"成功<br>"
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "aspdb"	'直接使用数据源
	strSql1="delete from user where username='"&username&"'"
	conn.execute(strSql1)
	
	strSql2="select username,password,sex,email from user"
	'set rs=server.CreateObject("ADODB.Recordset")
	set rs=conn.execute(strSql2)
	do while not rs.eof
	%>
		<tr>
		<td><%=rs("username")%></td><td><%=rs("password")%></td>
		<td><%=rs("sex")%></td><td><%=rs("email")%></td>
		</tr>
	<%
		rs.MoveNext
	loop
	'end if
%>
</table>
</center>
</body>
</html>