<html>
<head>
<title>第9题</title>
</head>
<script language="vbscript">
	function check_info()
		MsgBox("用户名不能为空")
		check_info=false
		
	end function
</script>
<body>
	<form name="form1" onSubmit="vbscript:return check_info" action="" method="post">
	用户名：<input type="text" name="username">
	密码：<input type="password" name="password">
	性别：<input type="radio" name="sex" value="男" checked>男&nbsp;&nbsp;
	<input type="radio" name="sex" value="女">女&nbsp;&nbsp;
	<input type="radio" name="sex" value="保密">保密&nbsp;&nbsp;
	e-mail:<input type="radio" name="email">
	<input type="submit" value="提交">
	<input type="reset" value="重置">
	</form>
	
</body>
</html>