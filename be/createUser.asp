 <%@LANGUAGE="VBSCRIPT" CODEPAGE="1255"%>
<!--#include file="condb.asp" -->
<!--#include file="common.asp" -->
<%

dim userName, userEmail, userPassword, relateToGroupId
userName = getFormValue("userName", true)
userEmail = getFormValue("userEmail", true)
userPassword = getFormValue("userPassword", true)
relateToGroupId = getFormValue("relateToGroupId", false)

dim uuid
Set TypeLib = CreateObject("Scriptlet.TypeLib")
uuid = Mid(TypeLib.Guid, 2, 36)

' insert

Set con = Server.CreateObject("ADODB.Connection")
con.Open AIcon_STRING 

sql_update = "INSERT INTO UserTable (userId, userName, userEmail, userPassword, relateToGroupId) VALUES ( '"&cstr(uuid)&"', '"&userName&"', '"&userEmail&"', '"&userPassword&"', '"&relateToGroupId&"' )"
con.Execute sql_update

con.Close
Set con = Nothing

' response json
Response.ContentType = "application/json"
response.Write("{""userId"": """&uuid&""", ""userName"":"""&URLDecode(userName)&""", ""userEmail"":"""&URLDecode(userEmail)&""", ""userPassword"":"""&URLDecode(userPassword)&""", ""relateToGroupId"":"""&URLDecode(relateToGroupId)&"""  }")

%>

