 <%@LANGUAGE="VBSCRIPT" CODEPAGE="1255"%>
<!--#include file="condb.asp" -->
<!--#include file="common.asp" -->
<%

dim groupName
groupName = Request.QueryString("groupName")
groupName = URLEncode(groupName)

dim uuid
Set TypeLib = CreateObject("Scriptlet.TypeLib")
uuid = Mid(TypeLib.Guid, 2, 36)

' insert to CHARGE TABLE

Set con = Server.CreateObject("ADODB.Connection")
con.Open AIcon_STRING 

sql_update = "INSERT INTO GroupTable (groupName, groupId) VALUES ('"&groupName&"', '"&cstr(uuid)&"')"
con.Execute sql_update

con.Close
Set con = Nothing

' response json
Response.ContentType = "application/json"
response.Write("{groupId: '"&uuid&"', groupName:'"&URLDecode(groupName)&"'}")

%>

