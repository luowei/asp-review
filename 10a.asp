<html>
<head>
<title>��10��</title>
</head>
<body>
<h1>��10��a</h1>
<script language="vbscript">
	function check_info()
		if form1.username.value="" then
			msgbox "�û�������Ϊ��"
			check_info=0
		elseif form1.password.value="" then
			msgbox "���벻��Ϊ��"
			check_info=0
		elseif form1.email.value="" then
			msgbox "email��ַ����Ϊ��"
			check_info=0
		else
			chck_info=1
		end if
	end function
</script>
<form name="form1" onSubmit="check_info" method="post" action="">
�û�����<input type="text" name="username"><br>
���룺<input type="password" name="password"><br>
�Ա�
<input type="radio" name="sex" value="��" checked>��&nbsp;&nbsp;
<input type="radio" name="sex" value="Ů">Ů<br>
email:<input type="text" name="email"><br>
<input type="submit" value="�ύ">&nbsp;&nbsp;
<input type="reset" value="����">
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
	response.Write "����user������д����"
	response.Write(username)
end if
%>

</body>
</html>