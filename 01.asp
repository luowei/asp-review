<html>
<head>
	<title>��1��</title>
</head>
<body>
	<%
	Response.Write Date() & "&nbsp;" & Time()
	Select Case Weekday(Date())
	case 1
		Response.Write "������"
	case 2
		Response.Write "����һ"
	case 3
		Response.Write "���ڶ�"
	case 4
		Response.Write "������"
	case 5
		Response.Write "������"
	case 6
		Response.Write "������"
	case 7
		Response.Write "������"
	End Select
	%>
</body>
</html>