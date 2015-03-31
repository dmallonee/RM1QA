<%@ Page Language="C#" Debug="false" AutoEventWireup="false" CodeFile="alerts_rate_management_car.aspx.cs" Inherits="alerts_rate_management_car" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 12.0"/>
<meta name="ProgId" content="FrontPage.Editor.Document"/>
<meta http-equiv="Content-Language" content="en-us"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252"/>
<title>Rate-Monitor by Rate-Highway, Inc. | Alerts! | Rate Management</title>
<script language="javascript" type="text/javascript" src="inc/sitewide.js"></script>
<script language="javascript" type="text/javascript">
function grid1_DoubleClick(item)
{
    window.open('rule_edit.asp?edit_mode=1&rate_rule_id=' + item.getMember('ID').get_text());
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
        var id = item.getMember('ID').get_text();
        
        switch(command)
        {
        case "Enable":
            window.open('enable_profiles_car.asp?profile_id=' + id);
            break;
        case "Disable":
            window.open('enable_profiles_car.asp?profile_id=' + id);
            break;
        case "Delete":
            window.open('search_profiles_maint_car.asp?profile_id=' + id);
            break;
        case "Exit":
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
          alert("item added");
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


function btnEnable_Click(grid, columnNumber)
{
    alert("testing 1-2-3");
    var id = GetCheckedItems(grid, columnNumber);
    alert(id);
    window.open('car_rate_rule_maint.asp?action=1&id=' + id);
}

function btnDisable_Click(grid, columnNumber)
{
    var id = GetCheckedItems(grid, columnNumber);
    window.open('car_rate_rule_maint.asp?action=1&id=' + id);
}

function btnDelete_Click(grid, columnNumber)
{
    var id = GetCheckedItems(grid, columnNumber);
    window.open('car_rate_rule_maint.asp?action=1&id=' + id);
}


</script>
<link rel="stylesheet" type="text/css" href="inc/rh_standard.css"/>
<link rel="stylesheet" type="text/css" href="inc/rh_report.css"/>
<link rel="stylesheet" type="text/css" href="inc/sitewideXXX.css"/>
<link rel="stylesheet" type="text/css" href="inc/GridStyle.css"/>
<link rel="stylesheet" type="text/css" href="inc/menuStyle.css" />
<link rel="stylesheet" type="text/css" href="inc/contextMenuStyle.css">
<link rel="stylesheet" type="text/css" href="inc/GridStyle.css">

</head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" onLoad="">
<form method="post" name="search_alerts" class="search" runat="server">
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
        <td><img src="images/h_alerts.gif" width="368" height="31"/></td>
        <td><map name="logout_map">
        <area alt="Click to log out of Rate-Monitor" href="http://www.rate-monitor.com/" shape="rect" coords="337, 10, 394, 25"/>
        </map>
        <img src="images/bottom_right.gif" width="402" height="31" usemap="#logout_map" border="0"/></td>
      </tr>
    </table>
    </td>
  </tr>
  <tr style="height:30px"><td>[<a href="alerts_rate_management_car.aspx?status=E" >Show Enabled Rules</a>] 
    [<a href="alerts_rate_management_car.aspx?status=A" >Show All Rules</a>]&nbsp;</td></tr>
  <tr>
    <td>
    <ComponentArt:Grid id="Grid1" 
         DataSourceID="SqlAlertsRulesSource"
         AllowHorizontalScrolling="false"
         RunningMode="Client" 
         CssClass="Grid" 
         ShowHeader="true"
         ShowSearchBox="true"
         SearchTextCssClass="GridSearchText"
         SearchOnKeyPress="true"
         HeaderCssClass="GridHeader" 
         FooterCssClass="GridFooter" 
         GroupByCssClass="GroupByCell"
         GroupByTextCssClass="GroupByText"
         PageSize="40" 
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
         ImagesBaseUrl="images" 
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
         Width="100%" Height="400" runat="server">
         <ClientEvents>
            <ContextMenu EventHandler="Grid1_onContextMenu" />
          </ClientEvents>

         <Levels>
           <ComponentArt:GridLevel
            DataKeyField="ID"
            HoverRowCssClass ="RowHover" 
            ShowSelectorCells="True" 
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
             <Columns>             
               <ComponentArt:GridColumn DataField="Description"  IsSearchable="True"  HeadingCellCssClass="FirstHeadingCell" DataCellCssClass="FirstDataCell"  HeadingText="Description"  AllowGrouping="False"  Width="300" />
               <ComponentArt:GridColumn AllowEditing="True" Visible="true" ColumnType="CheckBox" />
               <ComponentArt:GridColumn DataField="Comp Set"     IsSearchable="True"  HeadingCellCssClass="HeadingCell" DataCellCssClass="DataCell"  HeadingText="Comp. Set" Width="75" />
               <ComponentArt:GridColumn DataField="Car Type"     IsSearchable="True"  Width="40"  HeadingText="Car(s)"/>
               <ComponentArt:GridColumn DataField="Response"     IsSearchable="True"  Width="200" />
               <ComponentArt:GridColumn DataField="Response Amt" IsSearchable="True" Align="Right" Width="55" HeadingText="Rsp. Amt."  />
               <ComponentArt:GridColumn DataField="Rate Ceiling" IsSearchable="True" Align="Right" Width="55" HeadingText="Ceiling"/>
               <ComponentArt:GridColumn DataField="Rate Floor"   IsSearchable="True" Align="Right" Width="55" HeadingText="Floor"/>
               <ComponentArt:GridColumn DataField="Status"       IsSearchable="True" Width="50" />
               <ComponentArt:GridColumn DataField="ID" Visible="False" />
             </Columns>
             
             <ConditionalFormats>
               <ComponentArt:GridConditionalFormat ClientFilter="DataItem.GetMember('Status').Value=='disabled'" RowCssClass="GrayRow" SelectedRowCssClass="SelectedRow" HoverRowCssClass="GrayRowHover" SelectedHoverRowCssClass="" />
            </ConditionalFormats>
           </ComponentArt:GridLevel>
            
         </Levels>
         <ClientTemplates>
         

         <ComponentArt:ClientTemplate Id="LoadingFeedbackTemplate" runat="server">
          <table cellspacing="0" cellpadding="0" border="0">
          <tr>
            <td style="font-size:10px;">Loading...&nbsp;</td>
            <td><img src="images/spinner.gif" width="16" height="16" border="0"/></td>
          </tr>
          </table>
          </ComponentArt:ClientTemplate> 
          <ComponentArt:ClientTemplate Id="SliderTemplate" runat="server">
            <table class="SliderPopup" cellspacing="0" cellpadding="0" border="0">
            <tr>
              <td valign="top" style="padding:5px;">
              <table width="100%" cellspacing="0" cellpadding="0" border="0">
                <tr>
                <td width="25" align="center" valign="top" style="padding-top:3px;"></td>
                <td>
                <table cellspacing="0" cellpadding="1" border="0" style="width:255px;">
                <tr>
                  <td colspan="2" style="font-family:verdana;font-size:11px;font-weight:bold;"><div style="overflow:hidden;width:250px;"><nobr>## DataItem.GetMember('Description').Value ##</nobr></div></td>
                </tr>
                <!-- 
                <tr>
                  <td style="font-family:verdana;font-size:11px;"><div style="overflow:hidden;width:115px;"><nobr>by StartedBy</nobr></div></td>
                  <td style="font-family:verdana;font-size:11px;color:#918B7A;" align="right"><div style="overflow:hidden;width:135px;"><nobr>LastPostDate</nobr></div></td>
                </tr>
                -->
                <!--
                <tr>
                  <td style="font-family:verdana;font-size:11px;"><b>TotalViews</b> Views</td>
                  <td style="font-family:verdana;font-size:11px;" align="right"><b>Replies</b> Replies</td>
                </tr>
                -->
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
                Alert <b>## DataItem.Index + 1 ##</b> of <b>## Grid1.RecordCount ##</b>
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
         SiteMapXmlFile="RuleGridMenuData.xml"
         ExpandSlide="none"
         ExpandTransition="Fade"
         ExpandDelay="200"
         CollapseSlide="none"
         CollapseTransition="Fade"
         Orientation="Vertical"
         CssClass="MenuGroup"
         DefaultGroupCssClass="MenuGroup"
         DefaultItemLookID="DefaultItemLook"
         DefaultGroupItemSpacing="1"
         ImagesBaseUrl="images/"
         EnableViewState="false"
         ContextMenu="Simple"
         runat="server">
       <ItemLooks>
          <ComponentArt:ItemLook LookID="DefaultItemLook" CssClass="MenuItem" HoverCssClass="MenuItemHover" ExpandedCssClass="MenuItemHover" LeftIconWidth="20" LeftIconHeight="18" LabelPaddingLeft="10" LabelPaddingRight="10" LabelPaddingTop="3" LabelPaddingBottom="4" />
          <ComponentArt:ItemLook LookID="BreakItem" CssClass="MenuBreak" />
       </ItemLooks>
       </ComponentArt:Menu>  
    <asp:SqlDataSource ID="SqlAlertsRulesSource" runat="server" ConnectionString="<%$ ConnectionStrings:ProductionConnectionString %>"
        SelectCommand="car_rate_rule_select_grid" SelectCommandType="StoredProcedure" ProviderName="System.Data.SqlClient">
        <SelectParameters>
            <asp:Parameter DefaultValue="null" Name="rate_rule_id" Type="DBNull" />
            <asp:CookieParameter CookieName="rmuserid" DefaultValue="33" Name="user_id" Type="Int32" />
            <asp:Parameter DefaultValue="1" Name="rate_rule_type_cd" Type="Int32" />
            <asp:QueryStringParameter DefaultValue="E" Name="rule_status" QueryStringField="status" Type="String" />
        </SelectParameters>
    </asp:SqlDataSource>
    </td>
  </tr>
</table>   

<table >
  <tr>
      <td>
        <input type="button" value="Enable"  ID="btnEnable"  OnClick="btnEnable_Click(Grid1, 9);" style="width: 88px" class="rh_button" />&nbsp;
        <input type="button" value="Disable" ID="btnDisable" OnClick="btnDisable_Click(Grid1, 9);" style="width: 88px" class="rh_button" />&nbsp;
        <input type="button" value="Delete"  ID="btnDelete"  OnClick="btnDelete_Click(Grid1, 9);" style="width: 88px" class="rh_button" />
      </td>
  </tr>
</table>
</form>
<p>
    [<a href="rule_edit.asp?rate_rule_id=0" target="_blank" >New Rule</a>]
    [<a target="_blank" href="alerts_rate_management_export.asp">Download cross-reference</a>]
    [<a target="_blank" href="alerts_rate_management_export_worksheet.asp">Download rule worksheet for upload</a>]
    [<a target="_blank" href="rule_upload.asp">Upload rules</a>]</p>
<!--#INCLUDE FILE="footer.asp"-->
<font size="2" >
<p align="center">&nbsp;</p>
</font>

        <div id="calbox" class="calboxoff"></div>


                         
<script language="javascript"></script>
</body>
</html>

<!--<script language="javascript">
	document.search_alerts.alert_desc.focus();
</script>-->
<!-- 
             <Columns>             
               <ComponentArt:GridColumn DataField="Description" IsSearchable="True"  HeadingCellCssClass="FirstHeadingCell" DataCellCssClass="FirstDataCell"  HeadingText="Description"  AllowGrouping="False"  Width="200" />
               <ComponentArt:GridColumn DataField="Rate Code" IsSearchable="True" /> 
               <ComponentArt:GridColumn DataField="Comp Set" IsSearchable="True"  />
               <ComponentArt:GridColumn DataField="Car Type" IsSearchable="True"   />
               <ComponentArt:GridColumn DataField="Situation" IsSearchable="True"  /> 
               <ComponentArt:GridColumn DataField="Situation Amt" IsSearchable="True"   Align="Right"  />
               <ComponentArt:GridColumn DataField="Response" IsSearchable="True"   />
               <ComponentArt:GridColumn DataField="Response Amt" IsSearchable="True"   Align="Right"   />
               <ComponentArt:GridColumn DataField="Rate Ceiling" IsSearchable="True"   Align="Right"   />
               <ComponentArt:GridColumn DataField="Rate Floor"  IsSearchable="True"  Align="Right"  />
               <ComponentArt:GridColumn DataField="Status"  IsSearchable="True"  />
               <ComponentArt:GridColumn DataField="ID" Visible="False" />
             </Columns>
 -->
 <script type="text/javascript">
    // Preload CSS-referenced images
    (new Image()).src = 'images/group_background.gif';
    (new Image()).src = 'images/break_bg.gif';
 </script>



