<html>
<head>
	<title>��3��</title>
</head>
<body>
	<script type="text/vbscript">
		function gotoPage()
			dim site
			site=siteLink.value
			if site="����" then
				window.location="http://www.sina.com"
			elseif site="����" then
				window.location="http://www.163.com"
			elseif site="�Ѻ�" then
				window.location="http://www.sohu.com"
			else site="�ٶ�"
				window.location="http://www.baidu.com"
			end if
		end function
	</script>
	<select name="siteLink" size="1" onChange="gotoPage()">
		<option value="����">����</option>
		<option value="����">����</option>
		<option value="�Ѻ�">�Ѻ�</option>
		<option value="�ٶ�">�ٶ�</option>
	</select>
</body>
</html>