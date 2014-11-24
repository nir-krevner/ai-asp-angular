<%
Function URLDecode(str) 
        str = Replace(str, "+", " ") 
        For i = 1 To Len(str) 
            sT = Mid(str, i, 1) 
            If sT = "%" Then 
                If i+2 < Len(str) Then 
                    sR = sR & _ 
                        Chr(CLng("&H" & Mid(str, i+1, 2))) 
                    i = i+2 
                End If 
            Else 
                sR = sR & sT 
            End If 
        Next 
        URLDecode = sR 
    End Function 
 
Function URLEncode(str) 
	URLEncode = Server.URLEncode(str) 
End Function 

function getFormValue(fieldName, encode)
	dim val
	val = Request.Form(fieldName)
	if (encode = true) then
		val = URLEncode(val)
	end if
	'return value
	getFormValue = val
end function 

function responseError(errType, errMsg)
	
	Select Case errType
		Case "validation"
			response.Status="400 Bad Request"	
		
	End Select
	
	response.Write(response.Status & " " & errMsg)
	response.End 

end function 



%>