<html>
<head>
<title>��9��</title>
</head>
<script language="vbscript">
	function check_info()
		if form1.username.value="" then
			msgbox("�û�������Ϊ��")
			check_info=0
		elseif form1.password.value="" then
			msgbox("���벻��Ϊ��")
			check_info=0
		elseif form1.email.value="" then
			msgbox("email����Ϊ��")
			check_info=0
		else
		end if
	end function
</script>
<body>
<h1>��9��a</h1>
<h2>�û�ע��</h2>
	<form name="form1" onSubmit="check_info" action="" method="post">
	�û�����<input type="text" name="username"><br>
	���룺<input type="password" name="password"><br>
	�Ա�<input type="radio" name="sex" value="��" checked>��&nbsp;&nbsp;
	<input type="radio" name="sex" value="Ů">Ů<br>
	e-mail:<input type="text" name="email"><br>
	<input type="submit" value="�ύ">
	<input type="reset" value="����">
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