 <%@LANGUAGE="VBSCRIPT" CODEPAGE="1255"%>
<!--#include file="condb.asp" -->
<!--#include file="common.asp" -->
<%

dim groupId
groupId = Request.QueryString("groupId")

' get
Set groupRec = Server.CreateObject("ADODB.Recordset")
groupRec.ActiveConnection = AIcon_STRING
groupRec.Source = "SELECT * FROM GroupTable WHERE groupId='"&groupId&"'"
groupRec.CursorType = 2
groupRec.CursorLocation = 2
groupRec.LockType = 2
groupRec.Open()
	
if (groupRec.eof) then
	response.Write("eof")
else 
	' response json
	Response.ContentType = "application/json"
	response.Write("{groupId: '"&groupRec.fields.item("groupId").value&"', groupName:'"&URLDecode(groupRec.fields.item("groupName").value)&"'}")
end if

groupRec.Close
Set groupRec = Nothing
%>

