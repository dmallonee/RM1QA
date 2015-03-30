<%@ Page Language="C#" AutoEventWireup="true" CodeFile="~/search_queue_car.aspx.cs" Inherits="search_queue_car" Debug="true" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Queue</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<script language="Javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="Javascript" type="text/javascript" src="inc/header_menu_support.js" ></script>
<script language="Javascript" type="text/javascript" src="inc/js_calendar_v2.js"></script>

<script type='text/javascript' language="Javascript">
function grid1_DoubleClick(item)
{
    window.open('car_report_by_type.asp?reportrequestid=XXX&security_code=XXX');
}
function Grid1_onContextMenu(sender, eventArgs) 
     {
       Grid1.select(eventArgs.get_item()); 
       GridContextMenu.showContextMenu(eventArgs.get_event()); 
       GridContextMenu.set_contextData(eventArgs.get_item());     
     }
     function onContextMenuSelect()
     {
        window.open('car_report_by_type.asp?reportrequestid=XXX&security_code=XXX');
     }
</script>


<style>

.profile_header { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:0; background-color:#879AA2; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_light { height="68" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#CFD7DB; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
.profile_dark { height="48" text-align:left; padding-left:3; padding-right:3; padding-top:3; background-color:#C0C0C0; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10pt; vertical-align:bottom; text-align:left }
div#ExtraDay { margin: 0px 20px 0px 20px; display: none; }


</style>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewide.css">
<link rel="stylesheet" type="text/css" href="inc/GridStyle.css">
<link rel="stylesheet" type="text/css" href="inc/contextMenuStyle.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >

	
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<form id="Form1" method="post" name="search_alerts" class="search" runat="server">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/top_tile.gif" style="width: 1496px">
    <a target="_blank" href="http://www.rate-highway.com">
    <img src="images/top.jpg" width="770" height="91" border="0" ></a></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/b_tile.gif">

    <table width="400" border="0" cellspacing="0" cellpadding="0" id="page_header_buttons"><tr><td><img src="images/b_left.jpg" width="62" height="32"></td><td> <a href="search_profiles_car.aspx" onmouseover="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td><td> <a href="search_queue_car.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td><td> <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td><td> <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td><td> 
		<a href="alerts_rate_management_car.aspx"> <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td><td> <a href="javascript:not_enabled()" > <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td><td> <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td></tr></table>
    </td>

  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/med_bar_tile.gif">
    <img src="images/med_bar.gif" width="12" height="8"></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/user_tile.gif">

    <table width="100" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td><img src="images/user_left.gif" width="580" height="31"></td>
        <td background="images/user_tile.gif">
        <table width="100" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td valign="bottom">
            <table width="100" border="0" cellspacing="0" cellpadding="0">
              <tr>

                <td align="right">
                <div align="right">
                  <font face="Vendana, Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">
                  User: <%= UserName %></font></div>
                </td>
              </tr>
              <tr>
                <td><img src="images/separator.gif" width="183" height="6"></td>

              </tr>
            </table>
            </td>
            <td><img src="images/user_tile.gif" width="7" height="31"></td>
          </tr>
        </table>
        </td>
      </tr>
    </table>

    </td>
  </tr>
</table>


<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="images/h_tile.gif">
    <table width="100" border="0" cellspacing="0" cellpadding="0" id="table1">
      <tr>
        <td><img src="images/h_search_que.gif" width="368" height="31"></td>
        <td><map name="logout_map">

        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr style="height:30px"><td>&nbsp;</td></tr>
  <tr>
    <td>
    <ComponentArt:Grid id="Grid1" 
         DataSourceID="CarShopRequestSource"
         RunningMode="Client" 
         CssClass="Grid" 
         ShowHeader="true"
         ShowSearchBox="true"
         SearchTextCssClass="GridHeaderText"
         SearchOnKeyPress="true"
         HeaderCssClass="GridHeader" 
         FooterCssClass="GridFooter" 
         GroupByCssClass="GroupByCell"
         GroupByTextCssClass="GroupByText"
         PageSize="25" 
         PagerStyle="Slider" 
         PagerTextCssClass="GridFooterText"
         PagerButtonWidth="41"
         PagerButtonHeight="22"
         SliderHeight="20"
         SliderWidth="150" 
         SliderGripWidth="9" 
         SliderPopupOffsetX="20"
         SliderPopupClientTemplateId="SliderTemplate" 
         GroupingPageSize="5"
         PreExpandOnGroup="true"
         ImagesBaseUrl="images/" 
         PagerImagesFolderUrl="images/pager/"
         TreeLineImagesFolderUrl="images/lines/" 
         TreeLineImageWidth="22" 
         TreeLineImageHeight="19" 
         IndentCellWidth="22" 
         GroupingNotificationTextCssClass="GridHeaderText"
         GroupBySortAscendingImageUrl="group_asc.gif"
         GroupBySortDescendingImageUrl="group_desc.gif"
         GroupBySortImageWidth="10"
         GroupBySortImageHeight="10"
         LoadingPanelClientTemplateId="LoadingFeedbackTemplate"
         LoadingPanelPosition="MiddleCenter"
         ClientSideOnDoubleClick="grid1_DoubleClick"
         EditOnClickSelectedItem = "false"
         Width="100%" Height="600" runat="server">
         <ClientEvents>
            <ContextMenu EventHandler="Grid1_onContextMenu" />
          </ClientEvents>
         <Levels>
           <ComponentArt:GridLevel
             DataKeyField="shop_request_id"
             HoverRowCssClass ="RowHover"
            ShowTableHeading="false" 
            ShowSelectorCells="true" 
            SelectorCellCssClass="SelectorCell"
            SelectorCellWidth="18"
            SelectorImageUrl="selector.gif"
            SelectorImageWidth="17"
            SelectorImageHeight="15"
            HeadingSelectorCellCssClass="SelectorCell" 
            HeadingCellCssClass="HeadingCell" 
            HeadingRowCssClass="HeadingRow" 
            HeadingTextCssClass="HeadingCellText"            
            DataCellCssClass="DataCell" 
            RowCssClass="Row" 
            SelectedRowCssClass="SelectedRow"
            SortAscendingImageUrl="asc.gif" 
            SortDescendingImageUrl="desc.gif" 
            SortImageWidth="10" 
            SortImageHeight="10" >
            <ConditionalFormats>
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('request_status').Value=='Running'" RowCssClass="BlueRow" SelectedRowCssClass="SelectedRow" />
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('request_status').Value=='Cancelled'" RowCssClass="GrayRow" SelectedRowCssClass="SelectedRow" />
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('request_status').Value=='New'" RowCssClass="GreenRow" SelectedRowCssClass="SelectedRow" />
            </ConditionalFormats>
             <Columns>             
               <ComponentArt:GridColumn DataField="shop_request_id" Align="Left" IsSearchable="true"  HeadingCellCssClass="FirstHeadingCell" DataCellCssClass="FirstDataCell"  HeadingText="Search ID"  AllowGrouping="false"  Width="100" FixedWidth="False" />
               <ComponentArt:GridColumn DataField="request_status"  HeadingText="Status" IsSearchable="true"   Align="Left" />
               <ComponentArt:GridColumn DataField="scheduled_dttm"  HeadingText="Request (pst)"  FormatString="mm/dd/yyyy"  IsSearchable="true"   Align="Left"  />
               <ComponentArt:GridColumn DataField="email_dttm"  HeadingText="Email Date" IsSearchable="true" Visible="false" Align="Left"   />
               <ComponentArt:GridColumn DataField="alert_dttm"  HeadingText="Alert Date" IsSearchable="true" Visible="false"   Align="Left"  /> 
               <ComponentArt:GridColumn DataField="client_userid"  HeadingText="User Name" IsSearchable="true"   Align="Right"  />
               <ComponentArt:GridColumn DataField="profile_desc"  HeadingText="Profile" IsSearchable="true"   Align="Left"   />
               <ComponentArt:GridColumn DataField="data_sources"  HeadingText="Source" IsSearchable="true"   Align="Right"   />
               <ComponentArt:GridColumn DataField="work_units"  HeadingText="Rates Expected" IsSearchable="true"   Align="Right"   />
               <ComponentArt:GridColumn DataField="work_units_complete"   HeadingText="Rates Completed" IsSearchable="true"  Align="Right"  />
               <ComponentArt:GridColumn DataField="city_cd"  HeadingText="Pickup City"  IsSearchable="true"  Align="Left"  />
               <ComponentArt:GridColumn DataField="begin_arv_dt"   HeadingText="First Rental Date" FormatString="mm/dd/yyyy"   IsSearchable="true"  Align="Left"  />
               <ComponentArt:GridColumn DataField="end_arv_dt"   HeadingText="Last Rental Date" FormatString="mm/dd/yyyy"   IsSearchable="true"  Align="Left"  />
               <ComponentArt:GridColumn DataField="shop_car_type_cds"  HeadingText="Car Types" IsSearchable="true" Align="Left" />
               <ComponentArt:GridColumn DataField="vend_cd"  HeadingText="Companies" IsSearchable="true" Align="Left" />
             </Columns>
           </ComponentArt:GridLevel>
         </Levels>
         <ClientTemplates>
         

         <ComponentArt:ClientTemplate Id="LoadingFeedbackTemplate">
          <table cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td style="font-size:10px;">Loading...&nbsp;</td>
            <td><img src="images/spinner.gif" width="16" height="16" border="0"></td>
          </tr>
          </table>
          </ComponentArt:ClientTemplate> 
          <ComponentArt:ClientTemplate Id="SliderTemplate">
            <table class="SliderPopup" cellspacing="0" cellpadding="0" border="0">
            <tr>
              <td valign="top" style="padding:5px;">
              <table width="100%" cellspacing="0" cellpadding="0" border="0">
              <tr>
                <td width="25" align="center" valign="top" style="padding-top:3px;"></td>
                <td>
                <table cellspacing="0" cellpadding="2" border="0" style="width:255px;">
                <!--<tr>
                  <td style="font-family:verdana;font-size:11px;"><div style="overflow:hidden;width:115px;"><nobr>## DataItem.GetMember('StartedBy').Value ##</nobr></div></td>
                  <td style="font-family:verdana;font-size:11px;"><div style="overflow:hidden;width:135px;"><nobr>## DataItem.GetMember('LastPostDate').Text ##</nobr></div></td>
                </tr>-->
                <tr>
                  <td colspan="2">
                  <table cellspacing="0" cellpadding="0" border="0" width="100%">
                  <tr>
                    <td width="230" colspan="2" style="font-family:verdana;font-size:11px;font-weight:bold;"><div style="text-overflow:ellipsis;overflow:hidden;width:250px;"><nobr>## DataItem.GetMember('shop_request_id').Text ##</nobr></div></td>
                  </tr>
                  </table>                                    
                </tr>
                </table>    
                </td>
              </tr>
              </table>  
              </td>
              <td style="background-color:#CAC6D4;padding:2px;" align="center"></td>
            </tr>
            <tr>
              <td colspan="2" style="height:14px;background-color:#757598;">
              <table width="100%" cellspacing="0" cellpadding="0" border="0">
              <tr>
                <td style="padding-left:5px;color:white;font-family:verdana;font-size:10px;">
                Page <b>## DataItem.PageIndex + 1 ##</b> of <b>## Grid1.PageCount ##</b>

                </td>
                <td style="padding-right:5px;color:white;font-family:verdana;font-size:10px;" align="right">
                Profile <b>## DataItem.Index + 1 ##</b> of <b>## Grid1.RecordCount ##</b>
                </td>
              </tr>
              </table>  
              </td>
            </tr>
            </table>
          </ComponentArt:ClientTemplate>
         </ClientTemplates>

             
      </ComponentArt:Grid>
    <ComponentArt:Menu id="GridContextMenu" 
         SiteMapXmlFile="menuData.xml"
         ExpandSlide="none"
         ExpandTransition="fade"
         ExpandDelay="250"
         CollapseSlide="none"
         CollapseTransition="fade"
         Orientation="Vertical"
         CssClass="MenuGroup"
         DefaultGroupCssClass="MenuGroup"
         DefaultItemLookID="DefaultItemLook"
         DefaultGroupItemSpacing="1"
         ImagesBaseUrl="images/"
         EnableViewState="false"
         ContextMenu="Custom"
         runat="server">
       <ItemLooks>
         <ComponentArt:ItemLook LookID="DefaultItemLook" CssClass="MenuItem" HoverCssClass="MenuItemHover" LabelPaddingLeft="15" LabelPaddingRight="10" LabelPaddingTop="3" LabelPaddingBottom="3" />
         <ComponentArt:ItemLook LookID="BreakItem" ImageUrl="break.gif" CssClass="MenuBreak" ImageHeight="1" ImageWidth="100%" />
       </ItemLooks>
       </ComponentArt:Menu>  
    <asp:SqlDataSource ID="CarShopRequestSource" runat="server" ConnectionString="<%$ ConnectionStrings:developmentConnectionString %>"
        SelectCommand="car_shop_request_select1" SelectCommandType="StoredProcedure" CancelSelectOnNullParameter="false">
        <SelectParameters>
            <asp:Parameter Name="user_id" Type="String" />
            <asp:Parameter Name="user_role" Type="String" />
            <asp:Parameter Name="client_userid" Type="String" />
            <asp:Parameter Name="city_cd" Type="String" />
            <asp:Parameter  Name="shop_car_type_cds" Type="String" />
            <asp:Parameter  Name="vendor_cd" Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    </td>
  </tr>
  
</table>
<!-- JUSTTABS BOTTOM CLOSE -->
<p>&nbsp;</p>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp;</font></p>
<font size="2" >
<p align="center">&nbsp;</p>
</font>

        <div id="calbox" class="calboxoff"></div>
</form>
    </body>

</html>





