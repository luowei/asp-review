<html>
<head>
<title>��2��</title>
</head>
<body>
<%
	'����������£�
	function Rndab(a,b)
		Randomize Time()	'��ʼ������ӣ�����ÿ�β�����ֵ����һ����
		Rndab=Int(b-a+1)*Rnd()+a
	end function
	
	Response.Write "����һ��10��20֮����������"& Rndab(10,20)
%>
</body>
</html>