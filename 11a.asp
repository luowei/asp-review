<html>
<head>
<title>��11��</title>
</head>
<script language="vbscript">
	sub DelCheck(Byval username)
		password = inputbox("������ʾ��123456","����������","",700,600)
		if password="123456" then
			window.location="11b.asp?username="&username
		else
			msgbox("�������")
			window.location="11a.asp"
		end if
	end sub
</script>
<body>
<center>
<h1>��11��a</h1>
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
<tr><th>�û���</th><th>����</th><th>�Ա�</th><th>email</th><th>ɾ��</th></tr>
<%
do while not rs.eof
%>
	<tr>
	<td><%=rs("username")%></td>
	<td><%=rs("password")%></td>
	<td><%=rs("sex")%></td>
	<td><%=rs("email")%></td>
	<td><input type="button" value="ɾ��" onClick="DelCheck('<%=rs("username")%>')"></td>
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