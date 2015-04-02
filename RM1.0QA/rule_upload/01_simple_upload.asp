<%@ Language=VBScript %>
<%
	Set Upload = Server.CreateObject("Persits.Upload")
	Upload.IgnoreNoPost = True

	Rem fails when it sames within the website directories for some reason
	Count = Upload.Save("C:\rule_upload")
	'Count = Upload.Save("C:\Inetpub\wwwroot\rule_upload")

	rootDir = Server.MapPath("/")

	Response.Write Count & " files(s) uploaded:" & chr(13) & chr(10) & chr(13) & chr(10)

	For Each File in Upload.Files
		Response.Write File.FileName & chr(13) & chr(10)
	Next
%>


