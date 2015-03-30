<%@ language=vbscript %>
<%    

'	Option Explicit
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
	Response.Buffer = TRUE

	Dim objDOMDoc
	On Error Resume Next

	Set objDOMDoc = Server.CreateObject("Msxml2.DOMDocument.4.0")
	If Err Then
		Response.ContentType = "text/html"
		Response.Write "Error: " & Err.Description
		Err.Clear()
		Response.End
	End If
	objDOMDoc.async = False
	If objDOMDoc.load (Request) Then
		Response.ContentType = "text/xml"
		Response.write objDOMDoc.xml
	Else
		Response.ContentType = "text/html"
		Response.write "Error: " & objDOMDoc.parseError.reason
	End If	
	

sdata = Request.BinaryRead(Request.TotalBytes)
sdata = BinaryToString(sdata)
Response.binaryWrite sdata
Set fstemp = server.CreateObject("Scripting.FileSystemObject")
Set filetemp = fstemp.CreateTextFile(server.mappath(".") & "\myipaddresses.htm", true)
' true = file can be over-written if it exists
' false = file CANNOT be over-written if it exists
filetemp.write( sdata)
filetemp.Close
Function BinaryToString(Binary)
Dim cl1, cl2, cl3, pl1, pl2, pl3
Dim L
Dim tstChar
cl1 = 1
cl2 = 1
cl3 = 1
L = LenB(Binary)
Do While cl1<=L
tstChar=Chr(AscB(MidB(Binary,cl1,1)))
pl3 = pl3 & tstChar
cl1 = cl1 + 1
cl3 = cl3 + 1
If cl3>300 Then
pl2 = pl2 & pl3
pl3 = ""
cl3 = 1
cl2 = cl2 + 1
If cl2>200 Then
pl1 = pl1 & pl2
pl2 = ""
cl2 = 1
End If
End If
Loop
BinaryToString = pl1 & pl2 & pl3
End function

  If err.number <> 0 Then
	   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
	   response.write "<b>VBScript Errors Occured!<br>"
	   response.write "</b><br>"
	   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
	   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
	   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
	   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
	   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"

	End If

%>


