<html>
<head>
<meta http-equiv="refresh" content="3"> 
</head>
<body><center>
<%
	'��������ʾ���ͼƬ
	'���Ȳ���һ��10���ڵ��������
	Dim NumTemp
	Randomize Timer           '��������ڳ�ʼ��������ӣ�����ÿ�β�����ֵ��һ��
	NumTemp = Int((10 * Rnd()) + 1)
	'���ݸ�������ʾ��Ӧ��ͼƬ�������ϳ�����
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