<%
'功能：实现将满足用户条件的数据从相应的数据表中找出来，
'并以表格的形式显示到Web页面。（说明：要求用参数查询实现）
'author:罗维
'日期：2011-10-29
%>
<html>
<head>
<title>第12题</title>
</head>
<body>
<center>
<h1>第12题</h1>
<form name="form1" action="">
请输入查找的姓名：<input type="text" name="username">
<input type="submit" value="查找">
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
	prmType=200		'200表示变长字符串
	prmDirection=1
	prmSize=10
	prmValue=request.form("username")
	
	dim prm
	set prm=cmd.CreateParameter(prmName,prmType,prmDirection,prmSize,prmValue)
	cmd.Parameters.Append(prm)	'把prm参数对象添加到参数集
	
	dim rs
	cmd.CommandType=4	'表示查询指令是查询名
	cmd.CommandText="qryByUsername"
	Set rs=cmd.execute()
	
	do while not rs.eof
		response.Write "用户名：<b>"&rs("username") &"</b>&nbsp;"
		response.Write "密码：<b>"&rs("password") &"</b>&nbsp;"
		response.Write "性别：<b>"&rs("sex") &"</b>&nbsp;"
		response.Write "email：<b>"&rs("email") &"</b>&nbsp;"
		rs.moveNext
	loop
end if
%>
</center>
</body>
</html>