<html>
<head>
	<title>��6��</title>
</head>
<body>
<%
	if session("flat")<>1 then
		application.Lock()
		application("visitCount")=application("visitCount")+1
		Application.UnLock()
		session("flat")=1
	end if
	Dim visitCount
	visitCount=Application("visitCount")
	Response.Write "���ǵ�" & visitCount & "���ʱ�վ"
%>
</body>
</html>