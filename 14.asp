<html>
<head>
<title>第14题</title>
</head>
<body>
<h1>第14题</h1>
<html>
<BODY>
	<form action="" method="post">
		<%
		'以下连接数据库，建立一个Connection对象实例conn
		Dim conn,strConn 
		Set conn=Server.CreateObject("ADODB.Connection")
		strConn="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.Mappath("db.mdb")
		conn.Open strConn 
		'-----------------------------
		'以下建立Recordset对象实例rs，注意SQL语句中要用到上面获取的排序字段
		Dim rs,strSql
		Set rs=Server.CreateObject("ADODB.Recordset")
		'下面先查找出单选的题目
		strSql ="Select ID,Title,selectA,selectB,selectC,selectD,rightAnswer,type From Test where type='单选' Order By ID "
		rs.Open strSql,conn,1									'注意参数设置为键盘指针
		%>
		<h3>一、单选题</h3>
		<%
		Dim i  'i表示题目数
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

		<h3>二、多选题</h3>
		<%
		 '下面先查找出多选的题目
		 strSql ="Select ID,Title,selectA,selectB,selectC,selectD,rightAnswer,type From Test where type='多选' Order By ID "
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
		请输入您的姓名<input type="text" name="userName" >&nbsp;&nbsp;<input type="submit" value="交卷">
	</form>
	<%
	'如果提交了表单，就执行下面的判断语句
	If Trim(Request.Form("userName"))<>"" Then      
		Dim Grade    '分数 （单选答案：BCACD；多选答案：CD\ABCD\AD\B\ACD）
		Grade=0
		Dim j
		j=1  
		While j<i    'j表示当前是第几道题目，i表示总题目数。
			If Request.Form("q"&j)=Request.Form("Answer"&j) Then
				Grade=Grade+10
			End If
			j=j+1
		Wend
		
		'下面把该人的分数存到成绩表中。
		strSql="insert into score(userName,score) values('"&Request.Form("userName")&"',"&Grade&")"
		conn.execute(strSql)

		Response.Write "<font color='red'>您的分数为" & Grade & "</font>"
	End If
	set rs=nothing
	conn.close:set conn=nothing
	%>
</BODY>
</HTML>
</body>
</html>