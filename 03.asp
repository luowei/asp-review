<html>
<head>
	<title>第3题</title>
</head>
<body>
	<script type="text/vbscript">
		function gotoPage()
			dim site
			site=siteLink.value
			if site="新浪" then
				window.location="http://www.sina.com"
			elseif site="网易" then
				window.location="http://www.163.com"
			elseif site="搜狐" then
				window.location="http://www.sohu.com"
			else site="百度"
				window.location="http://www.baidu.com"
			end if
		end function
	</script>
	<select name="siteLink" size="1" onChange="gotoPage()">
		<option value="新浪">新浪</option>
		<option value="网易">网易</option>
		<option value="搜狐">搜狐</option>
		<option value="百度">百度</option>
	</select>
</body>
</html>