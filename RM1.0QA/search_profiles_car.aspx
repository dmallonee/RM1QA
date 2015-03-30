<%@ Page Language="C#" AutoEventWireup="false" CodeFile="~/search_profiles_car.aspx.cs" Inherits="search_profiles_car" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Language" content="en-us">
<title>Rate-Monitor by Rate-Highway, Inc. | Search Profiles</title>
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
     function onContextMenuSelect(command)
     {
        var item = GridContextMenu.get_contextData();
        var id = item.getMember('profile_id').get_text();
        
        switch(command)
        {
        case "Run":        
            window.open('run_profiles_car.asp?profile_id=' + id);
            break;
        case "Edit":
            window.location.href = "search_criteria_car.asp?profile=" + id;
            break;
        case "Enable":
            window.open('enable_profiles_car.asp?profile_id=' + id);
            break;
        case "Disable":
            window.open('enable_profiles_car.asp?profile_id=' + id);
            break;
        case "Delete":
            window.open('search_profiles_maint_car.asp?profile_id=' + id);
            break;
        default:
            break;
        }
     }
     
     function GetCheckedItems(grid, columnNumber)
    {
      var checkedItems = new Array();
      var gridItem;
      var itemIndex = 0;
      var id = "";
      
      while(gridItem = grid.get_table().getRow(itemIndex))
      {
        if(gridItem.get_cells()[columnNumber].get_value())
        {
          checkedItems[checkedItems.length] = gridItem;
        }
        
        itemIndex++;
      }
      
        for (i = 0; i< checkedItems.length; i++)
        {
           id += checkedItems[i].getMember('profile_id').get_text() + ",";
        }
        id = id.substring(0, id.length - 1);
      return id;
    }

     function btnRun_Click(grid, columnNumber)
     {
        var id = GetCheckedItems(grid, columnNumber);       
        window.open('run_profiles_car.asp?profile_id=' + id);

     }
     
     function btnEnable_Click(grid, columnNumber)
     {
        var id = GetCheckedItems(grid, columnNumber);
        window.open('enable_profiles_car.asp?profile_id=' + id);
     }
     
     function btnDisable_Click(grid, columnNumber)
     {
        var id = GetCheckedItems(grid, columnNumber);
        window.open('enable_profiles_car.asp?profile_id=' + id);
     }
     
     function btnDelete_Click(grid, columnNumber)
     {
        var id = GetCheckedItems(grid, columnNumber);
        window.open('search_profiles_maint_car.asp?profile_id=' + id);
     }
</script>


<link rel="stylesheet" type="text/css" href="inc/rh_standard.css">
<link rel="stylesheet" type="text/css" href="inc/rh_report.css">
<link rel="stylesheet" type="text/css" href="inc/sitewideXXX.css">
<link rel="stylesheet" type="text/css" href="inc/contextMenuStyle.css">
<link rel="stylesheet" type="text/css" href="inc/css_calendar_v2.css" id="calendarcss" media="all" >
<link rel="stylesheet" type="text/css" href="inc/GridStyle.css">


	
</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onLoad="MM_preloadImages('images/b_search_pro_on.gif','images/b_search_que_on.gif','images/b_search_cri_on.gif','images/b_rate_on.gif','images/b_alert_on.gif','images/b_user_on.gif','images/b_system_on.gif')">
<form id="search_form" method="post" name="search_alerts" class="ca_grid" runat="server">
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

    <table width="400" border="0" cellspacing="0" cellpadding="0" id="page_header_buttons"><tr><td><img src="images/b_left.jpg" width="62" height="32"></td><td> <a href="search_profiles_car.asp" onmouseover="MM_swapImage('s1','','images/b_search_pro_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_pro_of.gif" name="s1" border="0" id="s1" width="183" height="32"></a></td><td> <a href="search_queue_car.asp" onmouseover="MM_swapImage('s2','','images/b_search_que_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_que_of.gif" name="s2" border="0" id="s2" width="97" height="32"></a></td><td> <a href="search_criteria_car.asp" onmouseover="MM_swapImage('s3','','images/b_search_cri_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_search_cri_of.gif" name="s3" border="0" id="s3" width="103" height="32"></a></td><td> <a href="javascript:not_enabled()" onmouseover="MM_swapImage('ra','','images/b_rate_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_rate_of.gif" name="ra" border="0" id="ra" width="88" height="32"></a></td><td> <a href="alerts_rate_management_car.aspx" onmouseover="MM_swapImage('al','','images/b_alert_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_alert_of.gif" name="al" border="0" id="al" width="53" height="32"></a></td><td> <a href="javascript:not_enabled()" > <img src="images/b_user_of.gif" name="us" border="0" id="us" width="126" height="32"></a></td><td> <a href="javascript:not_enabled()" onmouseover="MM_swapImage('sy','','images/b_system_on.gif',1)" onmouseout="MM_swapImgRestore()"> <img src="images/b_system_of.gif" name="sy" border="0" id="sy" width="58" height="32"></a></td></tr></table>
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
        <td><img src="images/h_search_profiles.gif" width="368" height="31"></td>
        <td><map name="logout_map">

        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25">
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr style="height:30px"><td>&nbsp;
  </td></tr>
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
            DataKeyField="profile_id"
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
            ColumnReorderIndicatorImageUrl="reorder.gif"
            SortImageWidth="10" 
            SortImageHeight="10" >
            <ConditionalFormats>
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('profile_status').Value=='Running'" RowCssClass="BlueRow" SelectedRowCssClass="SelectedRow" />
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('profile_status').Value=='Cancelled'" RowCssClass="GrayRow" SelectedRowCssClass="SelectedRow" />
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('profile_status').Value=='New'" RowCssClass="GreenRow" SelectedRowCssClass="SelectedRow" />
            </ConditionalFormats>
             <Columns> 
               
               <ComponentArt:GridColumn DataField="profile_id" Visible="False" />            
               <ComponentArt:GridColumn AllowEditing="True" Visible="true" ColumnType="CheckBox" />
               <ComponentArt:GridColumn DataField="profile_status" HeadingText="Status"        IsSearchable="true"   Align="Left" HeadingCellCssClass="FirstHeadingCell" DataCellCssClass="FirstDataCell"  AllowGrouping="true"  Width="70" FixedWidth="False" />
               <ComponentArt:GridColumn DataField="last_name"      HeadingText="User"          IsSearchable="true"   Align="Left"  Visible="true" />
               <ComponentArt:GridColumn DataField="desc"           HeadingText="Description"   IsSearchable="true"   Align="Left"  Visible="true"  Width="300" FixedWidth="False" />
               <ComponentArt:GridColumn DataField="city_cd"        HeadingText="City"          IsSearchable="true"   Align="Left"  Visible="true"  Width="100" FixedWidth="False" />
               <ComponentArt:GridColumn DataField="lor"            HeadingText="LOR"           IsSearchable="true"   Align="Right" Visible="true" /> 
               <ComponentArt:GridColumn DataField="days_out"       HeadingText="Days Out"      IsSearchable="true"   Align="Right" Visible="true" />
               <ComponentArt:GridColumn DataField="days_long"      HeadingText="Days"          IsSearchable="true"   Align="Right" Visible="true" />
               <ComponentArt:GridColumn DataField="shop_car_type_cds"  HeadingText="Car Types" IsSearchable="true"   Align="Left"  Visible="true"  Width="150" FixedWidth="False" />               
               <ComponentArt:GridColumn DataField="data_sources"   HeadingText="Data Sources"  IsSearchable="true"   Align="Left"  Visible="true" />
               <ComponentArt:GridColumn DataField="vend_cds"       HeadingText="Companies"     IsSearchable="true"   Align="Left"  Visible="true" />
             </Columns>
           </ComponentArt:GridLevel>
         </Levels>
         <ClientTemplates>
         <ComponentArt:ClientTemplate ID="CheckboxTemplate">
            <div>Test</div>
         </ComponentArt:ClientTemplate>
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
         SiteMapXmlFile="profileMenuData.xml"
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
         <ComponentArt:ItemLook LookID="BreakItem" ImageUrl="images/break.gif" CssClass="MenuBreak" ImageHeight="1" ImageWidth="100%" />
       </ItemLooks>
       </ComponentArt:Menu>  
    <asp:SqlDataSource ID="CarShopRequestSource" runat="server" ConnectionString="<%$ ConnectionStrings:developmentConnectionString %>"
        SelectCommand="car_shop_profile_select" SelectCommandType="StoredProcedure" CancelSelectOnNullParameter="False">
        <SelectParameters>
            <asp:Parameter Name="desc" Type="String" />
            <asp:Parameter Name="shop_car_type_cds" Type="String" DefaultValue="" />
            <asp:Parameter Name="vend_cds" Type="String" />
            <asp:CookieParameter CookieName="rmuserid" DefaultValue="" Name="user_id" Type="Int32" />
            <asp:Parameter Name="profile_id" Type="Int32" DefaultValue="" />
            <asp:Parameter Name="city_cd" Type="Int32" />
            <asp:Parameter DefaultValue="1" Name="enabled" Type="Int32" />
        </SelectParameters>      
    </asp:SqlDataSource>                  
    </td>                                  
  </tr> 
  <tr>
    <td style="height: 10px">
    <input type="button" value="Run" ID="lbtnRun"  OnClick="btnRun_Click(Grid1, 1);" />
    <input type="button" value="Enable" ID="btnEnable"  OnClick="btnEnable_Click(Grid1, 1);" />
    <input type="button" value="Disable" ID="btnDisable" OnClick="btnDisable_Click(Grid1, 1);" />
    <input type="button" value="Delete" ID="btnDelete" OnClick="btnDelete_Click(Grid1, 1);" /></td>
  </tr>                                
                                           
</table>                                   
</form>
<!-- JUSTTABS BOTTOM CLOSE -->
<p>&nbsp;</p>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">© 
2002 - 2013 - All rights reserved<br>
<b>Rate-Highway, Inc.</b><br>
</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">18001 Cowan, 
    Suite F<br />
    Irvine, CA 92614<br />
    (949) 614-0751&nbsp;</font></p>

        <div id="calbox" class="calboxoff"></div>

    </body>

</html>





