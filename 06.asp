<html>
<head>
	<title>第6题</title>
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
	Response.Write "您是第" & visitCount & "访问本站"
%>
</body>
</html>