<html>
<head>
<title>��11��</title>
</head>
<body>
<center>
<h1>��11��b</h1>
<table align="center" border="1">
<tr><th>�û���</th><th>����</th><th>�Ա�</th><th>email</th></tr>
<%
	'if isempty(request.Form("username")) then	
	dim conn,username
	username=request.QueryString("username")
	response.Write "ɾ��"&username&"�ɹ�<br>"
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "aspdb"	'ֱ��ʹ������Դ
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