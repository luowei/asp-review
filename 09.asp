<html>
<head>
<title>��9��</title>
</head>
<script language="vbscript">
	function check_info()
		MsgBox("�û�������Ϊ��")
		check_info=false
		
	end function
</script>
<body>
	<form name="form1" onSubmit="vbscript:return check_info" action="" method="post">
	�û�����<input type="text" name="username">
	���룺<input type="password" name="password">
	�Ա�<input type="radio" name="sex" value="��" checked>��&nbsp;&nbsp;
	<input type="radio" name="sex" value="Ů">Ů&nbsp;&nbsp;
	<input type="radio" name="sex" value="����">����&nbsp;&nbsp;
	e-mail:<input type="radio" name="email">
	<input type="submit" value="�ύ">
	<input type="reset" value="����">
	</form>
	
</body>
</html>