<html>
<head>
<title>第11题</title>
</head>
<script language="vbscript">
	sub DelCheck(Byval username)
		password = inputbox("密码提示：123456","请输入密码","",700,600)
		if password="123456" then
			window.location="11b.asp?username="&username
		else
			msgbox("密码错误")
			window.location="11a.asp"
		end if
	end sub
</script>
<body>
<center>
<h1>第11题a</h1>
<%
if isempty(request.Form("username")) then	
	dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "aspdb"
	strSql="select username,password,sex,email from user"
	set rs=server.CreateObject("ADODB.Recordset")
	rs.open strSql,conn,1,1
%>

<table align="center" border="1">
<tr><th>用户名</th><th>密码</th><th>性别</th><th>email</th><th>删除</th></tr>
<%
do while not rs.eof
%>
	<tr>
	<td><%=rs("username")%></td>
	<td><%=rs("password")%></td>
	<td><%=rs("sex")%></td>
	<td><%=rs("email")%></td>
	<td><input type="button" value="删除" onClick="DelCheck('<%=rs("username")%>')"></td>
	</tr>
<%
	rs.MoveNext
loop
end if
%>
</table>

</center>
</body>
</html>