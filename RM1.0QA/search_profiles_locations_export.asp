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
	
	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseServer

	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "org_profile_and_location_rpt"
	adoCmd.CommandType = adCmdStoredProc

	adoCmd.Parameters.Append adoCmd.CreateParameter("@org_id",              3, 1, 0)	
	adoCmd.Parameters("@org_id").Value = Session("org_id") 
	Set adoRS = adoCmd.Execute

	'--- Use the ExcelTemplate object's Open method to
	'--- open the template simpletemplate.xls.
		objSAXLTmplt.Open Server.MapPath(Application("vroot") & "templates/org_profile_location_template.xls")

	'--- Set the template's datasource to the Recordset 
	'--- returned from the database.  DataSource("Recordset") 
	'--- refers to the recordset specified by the template's 
	'--- data markers (%%=Recordset.ColumnName).
		objSAXLTmplt.DataSource("profiles") = adoRS 
		
	'--- Generate the spreadsheet, and open it in the browser.  
	'--- The Process method takes two parameters: the name and 
	'--- path of the generated spreadsheet, and an optional 
	'--- process method.  
		objSAXLTmplt.Process "Rate-Monitor search profiles.xls", saProcessOpenInExcel 'saProcessOpenInPlace 'saProcessOpenInExcel

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

