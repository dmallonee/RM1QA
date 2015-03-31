<%@ Language=VBScript %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>New Page 2</title>
</head>

<body>
  <%	For Each Whatever In Request.Form
						Response.Write Whatever & " = <b>" & Request.Form(Whatever) & "|</b> <br>"

       
					Next
				%>
<!--#INCLUDE FILE="footer.asp"-->
</body>
</html>
