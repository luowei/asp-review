<html>
<head>
<title>��16��</title>
</head>
<body>
<h1>��16��</h1>
<form name="form1" method="post" action="">
<p>
Ҫ����һ���ļ����д���һ��test1.txt�ļ�����д��5�����֣�<br>
�ٽ�test1.txt�еĵ�3������д��ڶ����ļ����д�����test2.txt
</p>
��ָ���ļ��еľ��Ի����·����<br>
��һ���ļ��У�<input type="text" name="path1" value=".\folder1"><br>
�ڶ����ļ��У�<input type="text" name="path2" value=".\folder2"><br>
<input type="hidden" name="flat" value="1">
<input type="submit" value="����"><br>
</form><br>
<%
if request.Form("flat") then
	dim fso
	set fso=server.CreateObject("Scripting.FileSystemObject")
	dim fld1,fld2
	fld1=server.MapPath(request.Form("path1"))
	response.Write(fld1&"<br>")
	fld2=server.MapPath(request.Form("path2"))
	fso.CreateFolder fld1
	fso.Createfolder fld2
	dim fle1,fle2,fle3
	set fle1=fso.CreateTextFile(fld1&"\test1.txt",true)
	set fle2=fso.CreateTextFile(fld2&"\test2.txt",true)
	fle1.writeline("���ǵ�һ��")
	fle1.writeline("���ǵڶ���")
	fle1.writeline("���ǵ�����")
	fle1.writeline("���ǵ�����")
	fle1.writeline("���ǵ�����")
	fle1.close
	response.Write("�Ѿ��ɹ�������һ���ļ�����д�����������֣�<br>")
	set fle3=fso.OpenTextFile(fld1&"\test1.txt",1,true)	'�ڶ�������1:ֻ����2:ֻд��3:��д������������true��ʾ���Ը���
	while fle3.Line<>3
		fle3.SkipLine
	wend
	fle2.writeline(fle3.readline)
	response.Write("�ѳɹ���ڶ����ļ�д���˵�һ���ļ��ĵ���������<br>")
	fle3.close
	fle2.close
end if
%>
</body>
</html>