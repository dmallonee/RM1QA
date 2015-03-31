<%@ Language=VBScript %>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<% Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

   Server.ScriptTimeout = 30

	On Error Resume Next

' Build connection string to aspgrid.mdb
strConnect = Session("pro_con")	

' Create an instance of AspGrid
Set Grid = Server.CreateObject("Persits.Grid")
'Grid.LoadParameters Server.MapPath("usergrid.xml")

' Connect to the database
Grid.Connect strConnect, "", ""

' Specify SQL statement
Grid.SQL = "select * from [user]"

' Hide identity column
Grid.Cols("user_id").Hidden = False

' Specify location of button images
Grid.ImagePath = "../images/aspgrid/"

' Properties
Grid.MaxRows = 25 
Grid.Table.Width = "1200"
Grid.MethodGet = False
Grid.CanEdit = False
Grid.CanDelete = False
Grid.CanAppend = False
Grid.Table.Align = "CENTER" 
 
Grid.Cols("user_id").Caption = "ID"
Grid.Cols("user_id").Header.Width = "50"
Grid.Cols("email_address").Caption = "Email"

Grid.Cols(1).Header.Font.Class = "table_header_first"
Grid.ColRange(2, 14).Header.Font.Class = "table_header"

Grid.ColRange(2, 14).Cell.AltBGColor = "#CFD7DB"
Grid.ColRange(2, 14).Cell.BGColor = "#B2BEC4"
Grid.ColRange(2, 14).Header.BGColor = "#879AA2"
Grid.ColRange(1, 14).Cell.Class = "cell_data"
Grid.ColRange(1, 14).CanSort = True 
Grid.ColRange(1,  1).Header.BGColor = "#E07D1A"
Grid.Cols(0).Footer.BGColor = "#879AA2"
Grid.Cols(0).Header.BGColor = "#879AA2"

Grid.ColRange(1, 1).Cell.AltBGColor = "#FDC677"
Grid.ColRange(1, 1).Cell.BGColor = "#FDC677"

Grid(1).Cell.Align = "CENTER"


Grid("expiration").FormatDate "%m/%d/%y"
Grid("last_login").FormatDate "%m/%d/%y" '"%b %d, %Y"
Grid("modified").FormatDate "%m/%d/%y"

Grid("lob_id").Array = Array("Air", "Car", "Hotel")
Grid("lob_id").VArray = Array(3, 2, 1)


Grid.Table.CellSpacing = 0
Grid.Table.CellPadding = 1 
Grid.Table.Caption = "Rate-Monitor Users"
Grid.Table.Caption.Font.Face = "Arial"
Grid.Table.Caption.Font.Bold = True

'Grid.LoadParameters Server.MapPath("table_user_grid.xml")

' Display grid
Grid.Display
Grid.Disconnect
Response.Write Grid.Expires

If err.number <> 0 Then
		   pad="&nbsp;&nbsp;&nbsp;&nbsp;"
		   response.write "<b>An error cccured while collecting transaction information</b><br>"
		   response.write pad & "Error Number= #<b>" & err.number & "</b><br>"
		   response.write pad & "Error Desc.= <b>" & err.description & "</b><br>"
		   response.write pad & "Help Context= <b>" & err.HelpContext & "</b><br>"
		   response.write pad & "Help File Path=<b>" & err.helpfile & "</b><br>"
		   response.write pad & "Error Source= <b>" & err.source & "</b><br><hr>"
	
End If


%> 
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252" />
<title>Untitled 1</title>
<link rel="stylesheet" type="text/css" href="aspgrid.css">
</head>

<body>

</body>

</html>
