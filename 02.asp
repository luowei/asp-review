<html>
<head>
<title>第2题</title>
</head>
<body>
<%
	'随机函数如下：
	function Rndab(a,b)
		Randomize Time()	'初始随机种子，否则每次产生的值都是一样的
		Rndab=Int(b-a+1)*Rnd()+a
	end function
	
	Response.Write "产生一个10到20之间的随机数："& Rndab(10,20)
%>
</body>
</html>