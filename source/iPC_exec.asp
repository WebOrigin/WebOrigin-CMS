<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="Connections/cms_conn_a.asp" -->
<%
text_area_a = Request.QueryString("editext")
id=Request.QueryString("id")
Title_A=Request.QueryString("Title_A")

Dim Update_Cn, StrSQL
Set Update_Cn = Server.CreateObject("ADODB.Connection")
Update_Cn.Open MM_cms_conn_a_STRING
StrSQL = "UPDATE Cont SET Text_Cont='" & text_area_a &"', Text_Title_Cont ='" & Title_A & "' WHERE Text_ID='" & id & "'"
Update_Cn.Execute StrSQL
Update_Cn.close
Set Update_Cn = Nothing
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>无标题文档</title>
</head>
<body>
Done!
</body>
</html>
