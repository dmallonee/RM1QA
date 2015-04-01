<%@ Language=VBScript %>
<% 'Option Explicit 
   Response.Expires = -1
   Response.cachecontrol="private" 
   Response.AddHeader "pragma", "no-cache"

'--- Copyright (c) 2002 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
%>
<!-- #INCLUDE FILE="inc/adovbs.asp" --> 
<!-- METADATA TYPE="TypeLib" UUID="{7BCD2133-64A0-4770-843C-090637114583}" -->
<%
'--- Script Setting
On Error Resume Next
'-----------------------------------------------------------------------
'--- Simple demonstration of the ExcelTemplate object at work
'--- binding to data markers in the original workbook
'---
'---
'--- Copyright (c) 2002 SoftArtisans, Inc.
'--- Mail: info@softartisans.com   http://www.softartisans.com
'-----------------------------------------------------------------------

	'--- Declarations
	Dim objSAXLTmplt
	Dim adoConnect
	Dim strConn, sqlText
	Dim RecordSet
	Dim adoCmd
	Dim adoRS
	Dim adoCmd2
	Dim adoRS2

	'--- Create an instance of the ExcelTemplate object.
		Set objSAXLTmplt = Server.CreateObject("SoftArtisans.ExcelTemplate") 

	'--- Create an ADO database connection, connect 
	'--- to the Northwind database, create a recordset 
	'--- of the top 20 records in the Orders table.
		'Set adoConnect = Server.CreateObject("ADODB.Connection")


	strProfileDesc    = Request.Query("profile_desc")
	strProfileCarType = Request.Query("profile_car_type")
	strProfileCarCo   = Request.Query("profile_car_co")
	strUserId         = Request.Cookies("rate-monitor.com")("user_id") 
	
	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseServer
	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shop_profile_select_export"
	adoCmd.CommandType = adCmdStoredProc
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@desc",              200, 1, 255)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cds", 200, 1, 4096)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_cds",          200, 1, 1024)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@user_id",             3, 1, 0)
	adoCmd.Parameters.Append adoCmd.CreateParameter("@profile_id",          3, 1, 0)

	adoCmd.Parameters("@user_id").Value = strUserId 

	If Trim(strProfileDesc) <> "" Then
		adoCmd.Parameters("@desc").Value = strProfileDesc 
	Else
		adoCmd.Parameters("@desc").Value = Null
	End If


	If Trim(strProfileCarType) <> "" Then
		adoCmd.Parameters("@shop_car_type_cds").Value = strProfileCarType 
	Else
		adoCmd.Parameters("@shop_car_type_cds").Value = Null
	End If
	

	If Trim(strProfileCarCo) <> "" Then
		adoCmd.Parameters("@vend_cds").Value = strProfileCarCo 
	Else
		adoCmd.Parameters("@vend_cds").Value = Null 
	End If

	Set adoRS = adoCmd.Execute

	'--- Use the ExcelTemplate object's Open method to
	'--- open the template simpletemplate.xls.
		objSAXLTmplt.Open Server.MapPath(Application("vroot") & "templates/profile_extract_template.xls")

	'--- Set the template's datasource to the Recordset 
	'--- returned from the database.  DataSource("Recordset") 
	'--- refers to the recordset specified by the template's 
	'--- data markers (%%=Recordset.ColumnName).
		objSAXLTmplt.DataSource("profiles") = adoRS 
		
	'--- Generate the spreadsheet, and open it in the browser.  
	'--- The Process method takes two parameters: the name and 
	'--- path of the generated spreadsheet, and an optional 
	'--- process method.  
		objSAXLTmplt.Process "Rate-Monitor search profiles for " &  adoRS.Fields("first_name").Value & " " & adoRS.Fields("last_name").Value & ".xls", saProcessOpenInExcel 'saProcessOpenInPlace 'saProcessOpenInExcel

	'--- Error Handling
	If Err.number <> 0 Then
		Response.Write "An error occured. The error description is '" & Err.Description & "' ." 
	Else
		'Response.Write "Bingo3!" 

		Response.end 
	End If

	Set objSAXLTmplt = Nothing
	Set adoRS = Nothing
	Set adoCmd = Nothing
	
%>

