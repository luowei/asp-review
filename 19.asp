<html>
<head>
<meta http-equiv="refresh" content="3"> 
</head>
<body><center>
<%
	'这段语句显示广告图片
	'首先产生一个10以内的随机整数
	Dim NumTemp
	Randomize Timer           '该语句用于初始化随机种子，否则每次产生的值都一样
	NumTemp = Int((10 * Rnd()) + 1)
	'根据该整数显示相应的图片，并加上超链接
	Select Case Numtemp
	Case 0,1,2,3
		Response.Write "<a href='http://blog.163.com/luowei505050@126/' target='_blank'><img width='300' height='70' border='2' src='advimages/163.png'></a>"
	Case 4,5,6
		Response.Write "<a href='http://www.cnblogs.com/luowei010101/' target='_blank'><img width='300' height='70' border='2' src='advimages/cnblogs.png'></a>"
	Case 7,8,9
		Response.Write "<a href='http://hi.baidu.com/' target='_blank'><img width='300' height='70' border='2' src='advimages/baidu.png'></a>"
	End Select	
%>
</center></body>
</html>