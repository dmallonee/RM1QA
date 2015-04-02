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

	strIPAddress = Request.Servervariables("REMOTE_ADDR") 

	intReportRequestId = Request("reportrequestid")
	
	strConn = Session("pro_con")
	
  	Set adoRS = CreateObject("ADODB.Recordset")
  	Set adoCmd = CreateObject("ADODB.Command")
	adoRS.CursorLocation = adUseServer
	
	adoCmd.ActiveConnection = strConn
	adoCmd.CommandText = "car_shopped_rate_select_rpt2"
	adoCmd.CommandType = adCmdStoredProc
	
	adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_request_id", 3, 1, 0, intReportRequestId) 'Request("reportrequestid"))
	adoCmd.Parameters.Append adoCmd.CreateParameter("@report", 3, 1, 0, 1)
	
	If Len(Request.QueryString("vend_override")) = 2 Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Request.QueryString("vend_override"))
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@vend_override", 200, 1, 2, Null)
	End If

	adoCmd.Parameters.Append adoCmd.CreateParameter("@ipaddress", 200, 1, 20, strIPAddress)
		
	If Request("car_type_cd") = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Null)	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@shop_car_type_cd", 200, 1, 4, Request("car_type_cd"))
	End If
	
	If Request("city_cd") = "" Then
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 5, Null)	
	Else
		adoCmd.Parameters.Append adoCmd.CreateParameter("@city_cd", 200, 1, 5, Request("city_cd"))
	End If

'	adoRS.Open adoCmd, , adOpenStatic, adLockReadOnly, adCmdStoredProc 
	Set adoRS = adoCmd.Execute

	Set adoRS = adoRS.NextRecordset

	'--- Use the ExcelTemplate object's Open method to
	'--- open the template simpletemplate.xls.
	'	objSAXLTmplt.Open Server.MapPath(Application("vroot") & "templates/simpletemplate.xls")
		objSAXLTmplt.Open Server.MapPath(Application("vroot") & "templates/standard_rate_report.xls")

'		objSAXLTmplt.Open Server.MapPath(Application("vroot") & "templates/" & Request.Form("excel_template"))

	'--- Set the template's datasource to the Recordset 
	'--- returned from the database.  DataSource("Recordset") 
	'--- refers to the recordset specified by the template's 
	'--- data markers (%%=Recordset.ColumnName).
		objSAXLTmplt.DataSource("Rates") = adoRS 
		
	'--- Generate the spreadsheet, and open it in the browser.  
	'--- The Process method takes two parameters: the name and 
	'--- path of the generated spreadsheet, and an optional 
	'--- process method.  
		objSAXLTmplt.Process "generic.xls", saProcessOpenInExcel 'saProcessOpenInPlace 'saProcessOpenInExcel

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

