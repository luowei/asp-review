<html>
<head>
	<title>第1题</title>
</head>
<body>
	<%
	Response.Write Date() & "&nbsp;" & Time()
	Select Case Weekday(Date())
	case 1
		Response.Write "星期日"
	case 2
		Response.Write "星期一"
	case 3
		Response.Write "星期二"
	case 4
		Response.Write "星期三"
	case 5
		Response.Write "星期四"
	case 6
		Response.Write "星期五"
	case 7
		Response.Write "星期六"
	End Select
	%>
</body>
</html>