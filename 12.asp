<%
'���ܣ�ʵ�ֽ������û����������ݴ���Ӧ�����ݱ����ҳ�����
'���Ա�����ʽ��ʾ��Webҳ�档��˵����Ҫ���ò�����ѯʵ�֣�
'author:��ά
'���ڣ�2011-10-29
%>
<html>
<head>
<title>��12��</title>
</head>
<body>
<center>
<h1>��12��</h1>
<form name="form1" action="">
��������ҵ�������<input type="text" name="username">
<input type="submit" value="����">
</form>
<%
if request.form("username")<>"" then
	dim conn,strConn
	set conn=server.CreateObject("ADODB.Connection")
	'strConn="Driver={Microsoft Access Driver (*.mdb)};Dbq="&server.MapPath("db.mdb")
	strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("db.mdb")
	conn.Open strConn
	
	dim cmd
	set cmd=server.CreateObject("ADODB.Command")
	cmd.ActiveConnection=conn
	
	dim prmName,prmType,prmDirection,prmSize,prmValue
	prmName="varUsername"
	prmType=200		'200��ʾ�䳤�ַ���
	prmDirection=1
	prmSize=10
	prmValue=request.form("username")
	
	dim prm
	set prm=cmd.CreateParameter(prmName,prmType,prmDirection,prmSize,prmValue)
	cmd.Parameters.Append(prm)	'��prm����������ӵ�������
	
	dim rs
	cmd.CommandType=4	'��ʾ��ѯָ���ǲ�ѯ��
	cmd.CommandText="qryByUsername"
	Set rs=cmd.execute()
	
	do while not rs.eof
		response.Write "�û�����<b>"&rs("username") &"</b>&nbsp;"
		response.Write "���룺<b>"&rs("password") &"</b>&nbsp;"
		response.Write "�Ա�<b>"&rs("sex") &"</b>&nbsp;"
		response.Write "email��<b>"&rs("email") &"</b>&nbsp;"
		rs.moveNext
	loop
end if
%>
</center>
</body>
</html>