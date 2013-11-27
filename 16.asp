<html>
<head>
<title>第16题</title>
</head>
<body>
<h1>第16题</h1>
<form name="form1" method="post" action="">
<p>
要求：向一个文件夹中创建一个test1.txt文件，并写入5行文字，<br>
再将test1.txt中的第3行文字写入第二个文件夹中创建的test2.txt
</p>
请指定文件夹的绝对或相对路径：<br>
第一个文件夹：<input type="text" name="path1" value=".\folder1"><br>
第二个文件夹：<input type="text" name="path2" value=".\folder2"><br>
<input type="hidden" name="flat" value="1">
<input type="submit" value="创建"><br>
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
	fle1.writeline("这是第一行")
	fle1.writeline("这是第二行")
	fle1.writeline("这是第三行")
	fle1.writeline("这是第四行")
	fle1.writeline("这是第五行")
	fle1.close
	response.Write("已经成功建立第一个文件，并写入了五行文字！<br>")
	set fle3=fso.OpenTextFile(fld1&"\test1.txt",1,true)	'第二个参数1:只读、2:只写、3:读写，第三个参数true表示可以覆盖
	while fle3.Line<>3
		fle3.SkipLine
	wend
	fle2.writeline(fle3.readline)
	response.Write("已成功向第二个文件写入了第一个文件的第三行文字<br>")
	fle3.close
	fle2.close
end if
%>
</body>
</html>