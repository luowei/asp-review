<html>
<head>
<title>��14��</title>
</head>
<body>
<h1>��14��</h1>
<html>
<BODY>
	<form action="" method="post">
		<%
		'�����������ݿ⣬����һ��Connection����ʵ��conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("db.mdb")
		conn.Open strConn 
		'-----------------------------
		'���½���Recordset����ʵ��rs��ע��SQL�����Ҫ�õ������ȡ�������ֶ�
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		'�����Ȳ��ҳ���ѡ����Ŀ
		strSql ="Select ID,Title,selectA,selectB,selectC,selectD,rightAnswer,type From Test where type='��ѡ' Order By ID "
		rs.Open strSql,conn,1									'ע���������Ϊ����ָ��
		%>
		<h3>һ����ѡ��</h3>
		<%
		Dim i  'i��ʾ��Ŀ��
		i=1
		while not rs.eof
			response.write rs("Title")
		%>
		<br><input type="radio" name="q<%=i%>" value="A">A.<%=Server.HtmlEncode(rs("selectA"))%>
		<br><input type="radio" name="q<%=i%>" value="B">B.<%=Server.HtmlEncode(rs("selectB"))%>
		<br><input type="radio" name="q<%=i%>" value="C">C.<%=Server.HtmlEncode(rs("selectC"))%>
		<br><input type="radio" name="q<%=i%>" value="D">D.<%=Server.HtmlEncode(rs("selectD"))%>
		<br><input type="hidden" name="Answer<%=i%>" value="<%=rs("rightAnswer")%>">
		<%
			i=i+1
			rs.moveNext
		 wend
		 rs.close
		%>

		<h3>������ѡ��</h3>
		<%
		 '�����Ȳ��ҳ���ѡ����Ŀ
		 strSql ="Select ID,Title,selectA,selectB,selectC,selectD,rightAnswer,type From Test where type='��ѡ' Order By ID "
		 rs.Open strSql,conn,1
		 while not rs.eof
			response.write rs("Title")
		%>
		<br><input type="checkbox" name="q<%=i%>" value="A">A.<%=Server.HtmlEncode(rs("selectA"))%>
		<br><input type="checkbox" name="q<%=i%>" value="B">B.<%=Server.HtmlEncode(rs("selectB"))%>
		<br><input type="checkbox" name="q<%=i%>" value="C">C.<%=Server.HtmlEncode(rs("selectC"))%>
		<br><input type="checkbox" name="q<%=i%>" value="D">D.<%=Server.HtmlEncode(rs("selectC"))%>
		<br><input type="hidden" name="Answer<%=i%>" value="<%=rs("rightAnswer")%>">
		<%
			i=i+1
			rs.moveNext
		 wend
		 rs.close
		%>
		<p>
		��������������<input type="text" name="userName" >&nbsp;&nbsp;<input type="submit" value="����">
	</form>
	<%
	'����ύ�˱�����ִ��������ж����
	If Trim(Request.Form("userName"))<>"" Then      
		Dim Grade    '���� ����ѡ�𰸣�BCACD����ѡ�𰸣�CD\ABCD\AD\B\ACD��
		Grade=0
		Dim j
		j=1  
		While j<i    'j��ʾ��ǰ�ǵڼ�����Ŀ��i��ʾ����Ŀ����
			If Request.Form("q"&j)=Request.Form("Answer"&j) Then
				Grade=Grade+10
			End If
			j=j+1
		Wend
		
		'����Ѹ��˵ķ����浽�ɼ����С�
		strSql="insert into score(userName,score) values('"&Request.Form("userName")&"',"&Grade&")"
		conn.execute(strSql)

		Response.Write "<font color='red'>���ķ���Ϊ" & Grade & "</font>"
	End If
	set rs=nothing
	conn.close:set conn=nothing
	%>
</BODY>
</HTML>
</body>
</html>