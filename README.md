# ASP-Review

## 项目简介
ASP (Active Server Pages) 机试考试复习题集项目。本项目包含了一系列经典的ASP编程练习题目，涵盖了ASP开发的核心知识点，包括表单处理、数据库操作、VBScript编程、页面跳转等常见Web开发场景。

## 技术栈
- **开发语言**: ASP (Active Server Pages)
- **脚本语言**: VBScript
- **数据库**: Microsoft Access (.mdb)
- **数据库连接**: ADO (ActiveX Data Objects)
- **Web服务器**: IIS (Internet Information Services)

## 功能特性
- 日期和时间显示
- 下拉菜单导航和页面跳转
- 表单数据验证 (客户端)
- 用户注册功能
- 数据库连接和操作 (增删改查)
- 广告图片管理
- 服务器端数据处理
- VBScript客户端脚本应用

## 项目文件说明

### 练习题文件
- **01.asp** - 日期时间显示，星期几判断
- **02.asp** - 基础页面示例
- **03.asp** - 下拉菜单选择，页面跳转 (新浪、网易、搜狐、百度)
- **05.asp** - 基础示例
- **06.asp** - 数据处理示例
- **07.asp** - 表单处理
- **09.asp** / **09a.asp** - 用户注册表单 (包含客户端验证)
- **09b.asp** - 用户注册结果显示
- **10a.asp** - 数据库查询和显示
- **10b.asp** - 数据库操作
- **11a.asp** - 用户登录
- **11b.asp** - 登录验证
- **12.asp** - 数据库增删改查综合应用
- **14.asp** - 广告图片管理
- **16.asp** - 数据处理
- **19.asp** - 其他功能示例

### 数据库文件
- **db.mdb** - Microsoft Access数据库文件，包含用户表等数据

### 其他文件
- **advimages/** - 广告图片资源目录
- **08计科ASP机试考试复习题.doc** - 考试复习题目文档

## 核心知识点

### 1. VBScript基础
```vbscript
' 日期时间处理
Response.Write Date() & "&nbsp;" & Time()

' 条件判断
Select Case Weekday(Date())
case 1
    Response.Write "星期日"
End Select
```

### 2. 表单处理
```vbscript
' 获取表单数据
username = request.Form("username")
password = request.Form("password")
```

### 3. 数据库操作
```vbscript
' ADO数据库连接
set conn = server.CreateObject("ADODB.Connection")
strConn = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("db.mdb")
conn.Open strConn

' SQL插入操作
strSql = "insert into user(username,password,sex,email) values(...)"
conn.execute(strSql)
```

### 4. 客户端验证
```vbscript
<script language="vbscript">
function check_info()
    if form1.username.value="" then
        msgbox("用户名不能为空")
        check_info=0
    end if
end function
</script>
```

### 5. 页面跳转
```vbscript
response.Redirect "09b.asp"
```

## 依赖要求
- **Web服务器**: IIS 5.0 及以上版本
- **操作系统**: Windows Server 2003/2008/2012 或 Windows XP/7/10 (带IIS组件)
- **数据库**: Microsoft Access 2003/2007 及以上版本
- **浏览器**: Internet Explorer 6+ (支持VBScript)

## 安装和运行方法

### 1. 配置IIS
```
控制面板 → 程序和功能 → 启用或关闭Windows功能 → Internet Information Services
```
勾选以下组件：
- Web管理工具
- 万维网服务
- ASP
- ISAPI扩展

### 2. 部署项目
1. 将项目文件复制到IIS网站目录 (如: C:\inetpub\wwwroot\asp-review)
2. 确保db.mdb文件有读写权限

### 3. 配置IIS站点
1. 打开IIS管理器
2. 右键"网站" → 添加网站
3. 设置网站名称、物理路径、端口
4. 确保"ASP"功能已启用

### 4. 访问项目
浏览器访问：
```
http://localhost/asp-review/01.asp
http://localhost/asp-review/09a.asp
```

## 数据库配置

### Access数据库连接字符串
```vbscript
' 方式1: OLEDB Provider
strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath("db.mdb")

' 方式2: ODBC Driver (推荐)
strConn = "Driver={Microsoft Access Driver (*.mdb)};Dbq=" & Server.Mappath("db.mdb")
```

### 数据库表结构
**user表** (用户信息)
- username - 用户名
- password - 密码
- sex - 性别
- email - 邮箱

## 学习要点
- ASP基础语法和页面结构
- VBScript服务器端和客户端编程
- Request对象的使用 (获取表单数据)
- Response对象的使用 (输出、重定向)
- Server对象的使用 (CreateObject、MapPath)
- ADO数据库访问技术
- SQL语句的编写
- 表单验证技术
- 会话管理基础

## 注意事项
1. **安全性**: 本项目为教学示例，未包含完整的安全措施 (如SQL注入防护、密码加密等)
2. **编码问题**: 文件可能使用GB2312编码，请确保IIS配置正确
3. **兼容性**: ASP是较老的技术，建议在学习目的下使用，生产环境推荐使用ASP.NET或其他现代框架
4. **数据库权限**: 确保IIS应用池账户对.mdb文件有读写权限

## 适用场景
- ASP课程学习和复习
- Web开发基础教学
- 机试考试准备
- 经典Web技术回顾

## 作者
luowei

## 相关文档
参考 **08计科ASP机试考试复习题.doc** 获取完整题目说明