         <tr  class="<%=strClass %>" >
          	<% strCarClasses = strCarClasses & adoRS2.Fields("car_type_cd").Value & "|" %>
            <td class="boxtitle" style="width: 14%"><font size="2"><%=adoRS2.Fields("car_type_cd").Value %></font></td>
            <td class="boxtitle" style="width: 14%">
            <% strDataValues = strDataValues & FormatNumber(((adoRS2.Fields("fleet_rsvd").Value + adoRS2.Fields("fleet_out").Value - adoRS2.Fields("fleet_invoiced").Value) / adoRS2.Fields("fleet_count").Value) * 100, 2) & "," %>
            <%=FormatPercent((adoRS2.Fields("fleet_rsvd").Value + adoRS2.Fields("fleet_out").Value - adoRS2.Fields("fleet_invoiced").Value) / adoRS2.Fields("fleet_count").Value) %></td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<a href="system_utilization_res_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>&no_show_percentage=<%=adoRS.Fields("no_show_percentage").Value %> "><%=adoRS2.Fields("fleet_rsvd").Value %></a></font></td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<a href="system_utilization_out_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>&no_show_percentage=<%=adoRS.Fields("no_show_percentage").Value %> "><%=adoRS2.Fields("fleet_out").Value %></a>&nbsp;&nbsp;&nbsp;</font>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="on_rent" size="20" readonly value="<%=adoRS2.Fields("fleet_out").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <td class="UtilGridValue" style="width: 14%"><font size="2"><%=adoRS2.Fields("fleet_canceled").Value %></font>
            <!--  
            <input type="text" name="canceled" size="20" readonly value="<%=adoRS2.Fields("fleet_canceled").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<a href="system_utilization_return_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>&fleet_invoiced=<%=adoRS2.Fields("fleet_invoiced").Value %> "><%=adoRS2.Fields("fleet_invoiced").Value %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="returned" size="20" readonly value="<%=adoRS2.Fields("fleet_invoiced").Value %>" style="text-align: right; width: 80px;">
            -->
            </td>
            <!-- 
            <td class="boxtitle" style="width: 14%">
            <input type="text" name="total" size="20" readonly value="<%=adoRS2.Fields("fleet_count").Value %>" style="text-align: right; width: 80px;">
            -->
            <td class="UtilGridValue" style="width: 14%"><font size="2">
			<a href="system_utilization_fleet_detail.asp?date=<%=Server.URLEncode(datUtilDate) %>&car_type_cd=<%=adoRS2.Fields("car_type_cd").Value %>&org_id=<%=adoRS.Fields("org_id").Value %>&city_cd=<%=strCityCd %>"><%=adoRS2.Fields("fleet_count").Value %></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
			</td>
          </tr>
          <% 	intCount = intCount + 1	%>
          <%   adoRS2.MoveNext         	%>
          <% Wend                     	%>
          <% End If 					%>
          <tr>
