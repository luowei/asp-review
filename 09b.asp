<html>
<head>
<title>��9��</title>
</head>
<body>
<h1>��9��b</h1>
<h2>��¼��֤</h2>
	<form name="form1" action="" method="post">
	�û�����<input type="text" name="username"><br>
	���룺<input type="password" name="password"><br>
	<input type="submit" value="��¼">
	<input type="reset" value="����">
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
		response.Write "<br>��¼ʧ��</body></html>"
		response.End()
	elseif rs("password")<>password then
		response.Write "<br>��¼ʧ��</body></html>"
		response.End()
	else
		session("username")=username
		response.Write "<br>��¼�ɹ�</body></html>"
	end if
else
end if
%>
</body>
</html>