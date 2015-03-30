<%@ Page Title="" Language="VB" MasterPageFile="~/popup.master" AutoEventWireup="false"
    CodeFile="threshold_maxmin_schedule_b.aspx.vb" Inherits="threshold_maxmin_schedule_b" %>
<%@ Register TagPrefix="ComponentArt" Namespace="ComponentArt.Web.UI" Assembly="ComponentArt.Web.UI" %>

<asp:Content ID="C1" ContentPlaceHolderID="cphMain" runat="Server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
</asp:ScriptManager>
<script language="javascript" type="text/javascript">
    function setChanges() {
        document.getElementById("currentState").value = 1;
    }
</script>
<div align="center">
   
 
  <table border="0" cellpadding="0" cellspacing="0" width="680">
    <tr>
        <td>
    <p align="left">
        <asp:HyperLink ID="linkBack" runat="server" Text="Select"></asp:HyperLink>
        &nbsp;&gt;&nbsp;Edit Schedule 
    </p>
    <p align="center" style="width: 800px;">
        <font size="5" color="#384F5B">Threshold Min &amp; Max Schedule</font>
        
        
    </p>
    </td>
    </tr>
    </table>
   
        <table border="0" cellpadding="0" cellspacing="0" width="680" align="center"><tr><td align="center">
        
<ComponentArt:TabStrip id="TabStrip1"
      DefaultGroupCssClass="TopGroup"
      DefaultItemLookId="TopLevelTabLook"
      DefaultSelectedItemLookId="SelectedTopLevelTabLook"
      DefaultChildSelectedItemLookId="SelectedTopLevelTabLook"
      DefaultGroupTabSpacing="0"
      AutoPostBackOnSelect="true"
	  ScrollingEnabled="true"
      ScrollLeftLookId="ScrollItem"
      ScrollRightLookId="ScrollItem"
      Width="678"
      runat="server">
    		
 		
    <ItemLooks>
	    <ComponentArt:ItemLook LookId="TopLevelTabLook" CssClass="TopLevelTab" HoverCssClass="TopLevelTabHover" LabelPaddingLeft="15" LabelPaddingRight="15" LabelPaddingTop="4" LabelPaddingBottom="4" />
		<ComponentArt:ItemLook LookId="SelectedTopLevelTabLook" CssClass="SelectedTopLevelTab" LabelPaddingLeft="15" LabelPaddingRight="15" LabelPaddingTop="4" LabelPaddingBottom="4"  />
		<ComponentArt:ItemLook LookId="Level2TabLook" CssClass="Level2Tab" HoverCssClass="Level2TabHover" LabelPaddingLeft="15" LabelPaddingRight="15" LabelPaddingTop="4" LabelPaddingBottom="4" />
		<ComponentArt:ItemLook LookId="SelectedLevel2TabLook" CssClass="SelectedLevel2Tab" LabelPaddingLeft="15" LabelPaddingRight="15" LabelPaddingTop="4" LabelPaddingBottom="4"  />
		<ComponentArt:ItemLook LookId="ScrollItem" CssClass="ScrollItem" HoverCssClass="ScrollItemHover" LabelPaddingLeft="5" LabelPaddingRight="5" LabelPaddingTop="0" LabelPaddingBottom="0" />
	</ItemLooks>
</ComponentArt:TabStrip>
    
   </td></tr></table>
        <asp:Label ID="lblCB" runat="server" Visible="false"></asp:Label>
    
    <asp:UpdatePanel ID="pnlHeader" runat="server">
        <ContentTemplate>
   
        
        <asp:label ID="lblStatus" runat="server" Visible="false"></asp:label>
    
        
           
                            
                            <table border="0" cellpadding="0" cellspacing="0" 
                class="xtan-border" width="680" align="center">
                                <tr>
                                    <td align="left" class="style3">
                                        <br />
                                    </td>
                                    <td style="width: 35%" align="left">
                                        &nbsp;</td>
                                </tr>
                               
                                <tr>
                                    <td align="center" class="style3" colspan="2">
                                        
                                        <font size="2">Schedule Name: </font>
                                    
                                        <asp:TextBox cssclass="textboxleft" ID="txtName" runat="server" Width="400px"></asp:TextBox>
                                    </td>
                                </tr>
                               
                                <tr>
                                    <td colspan="2" style="text-align: center">
                                        <br />
                                        <b><font size="2">Rate Amounts</font></b><br />
                                        (Empty cells mean no limit)<br />
                                        <br />
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" align="center">
                                        <asp:GridView ID="grdCarTypes" runat="server" DataKeyNames="city_cd" 
                                            AutoGenerateColumns="false" AllowSorting="True" width="400px" 
                                            CssClass="generic" Font-Names="Verdana" Font-Size="Small">
                                            <Columns>
                                                
                            <asp:TemplateField HeaderText="Car Type">
                                <ItemTemplate>
                                    <asp:Label ID="car_type_cd" width="85" runat="server" Text='<%# Bind("car_type_cd") %>' ></asp:Label>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Min">
                                <ItemTemplate>
                                    <asp:Textbox ID="min_amt" onChange="javascript:setChanges();" cssclass="textboxright" width="85" runat="server" Text='<%# Bind("min_amt", "{0:F2}") %>'></asp:Textbox>
                                </ItemTemplate>
                            </asp:TemplateField>
                             <asp:TemplateField HeaderText="Max">
                                <ItemTemplate>
                                    <asp:Textbox ID="max_amt" onChange="javascript:setChanges();" cssclass="textboxright" width="85" runat="server" Text='<%# Bind("max_amt", "{0:F2}") %>'></asp:Textbox>
                                </ItemTemplate>
                            </asp:TemplateField>
                                               
                                            </Columns>
                                        </asp:GridView>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" align="center">
                                        <br />
                                        <asp:Button ID="btnUpdate" CssClass="buttoncenter" runat="server" Text="Update" 
                                            Width="80px" />
                                        <asp:CheckBox ID="chkDebug" runat="server" Visible="false" /> 
                                        
                                        <br />
                                    </td>
                                </tr>
                            </table>
                            
                            </td></tr></table>
                        </div>
                 
                  
        </ContentTemplate>
    </asp:UpdatePanel>
    
    <asp:HiddenField ID="currentState" Value="0" runat="server" />
                  <asp:HiddenField ID="currentLocation" Value="-1" runat="server" />
                  <asp:HiddenField ID="currentMonth" Value="-1" runat="server" />
                  
               
</div>
</asp:Content>
