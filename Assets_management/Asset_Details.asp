<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%
ISD = Request.form("ISD")
IF ISD = "" Then
ISD = request.querystring("ISD")
End If
%>
<%Note = request.querystring("Note")
SType = Request.Form("SType")
Emp_Id = request.form("Emp_Id")
Release_justification = request.form("Release_justification")
   
    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1 or  Session("SUser_Privilege") = 2 or  Session("SUser_Privilege") = 3  Then
%>
<html dir="ltr" >
<script type="text/javascript">
    function showCustomer(str) {
        if (str == "") {
            document.getElementById("txtHint").innerHTML = "";
            return;
        }
        if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
            xmlhttp = new XMLHttpRequest();
        }
        else {// code for IE6, IE5
            xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
        }
        xmlhttp.onreadystatechange = function () {
            if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                document.getElementById("txtHint").innerHTML = xmlhttp.responseText;
            }
        }
        //alert("Relase_Assets.asp?q=" + str + "abc");
        xmlhttp.open("GET", "Relase_Assets.asp?q=" + str, true);
        xmlhttp.send();
    }
</script>
    <head>

        <title>Assets Details</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>	 
        <p>
            &nbsp;</p>
        <%	if Note = "Update_ok" then%><div class="note">Asset Released Successfully</div><% end if%>
		<br/><br/>
        <h1>Asset information</h1>
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = "SELECT * FROM  Assets where ISD_Name = '" & ISD & "'"
	    Set RS0 = con.execute(SQL0)

	    IF NOT RS0.EOF THEN
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM Assets_Type e WHERE e.Assets_Type_Id = '" & RS0("Assets_Type_Id") & "'"
	    Set RS1 = con.execute(SQL1)
        IF NOT RS1.EOF THEN
        %>&nbsp;
  <table >
   <tr><td>
    <table>
        <tr>
        <td class="cellT">ISD Name:</td><td class="cell"><%=RS0("ISD_Name")%></td>
        </tr>
        <tr>
        <td class="cellT">Type</td><td class="cell"><%=RS1("Assets_Type_Name")%></td>
        </tr>
        <tr>
        <td class="cellT">Assets Active</td><td class="cell">
        <%  IF RS0("Assets_Active") = 1     Then %>
                <font color="green">YES</font>
        <%  ELSEIF RS0("Assets_Active") = 0 Then %>
                <font color="Red">No</font>
        <%  Else %>
                UNKNOWN
        <%  End IF %>
        </td>
        </tr>
        <tr>
        <td class="cellT">Registration Date</td><td class="cell"><%=RS0("Date_Reg")%></td>
        </tr>
        <tr>
        <td class="cellT">Note</td><td class="cell"><%=RS0("Note")%></td>
        </tr>
	</table>
   </td>
   <td width="100">&nbsp;</td>
    <%  IF RS1("Assets_Type_Id") = 1 or RS1("Assets_Type_Id") = 2 Then %>
   <td valign=bottom>
        <table valign=bottom>
        <tr>
        <td class="cellT">Last Login By</td><td class="cell"><%=RS0("last_Name")%></td>
        </tr>
        <tr>
        <td class="cellT">Model</td><td class="cell"><%=RS0("Assets_Model")%></td>
        </tr>
        <tr>
        <td class="cellT">Processor Type</td><td class="cell"><%=RS0("Assets_Processor")%></td>
        </tr>
        <tr>
        <td class="cellT">Speed in GHz</td><td class="cell"><%=RS0("Assets_Speed")%></td>
        </tr>
        <tr>
        <td class="cellT">Current Ram</td><td class="cell"><%=RS0("Ram_Current")%></td>
        </tr>
        <tr>
        <td class="cellT">Max Ram</td><td class="cell"><%=RS0("Ram_Max")%></td>
        </tr>
        <tr>
        <td class="cellT">HD C used in MB</td><td class="cell"><%=RS0("HDC_Current")%></td>
        </tr>
        <tr>
        <td class="cellT">HD C Max in MB</td><td class="cell"><%=RS0("HDC_Max")%></td>
        </tr>
        <tr>
        <td class="cellT">HD D used in MB</td><td class="cell"><%=RS0("HDD_Current")%></td>
        </tr>
        <tr>
        <td class="cellT">HD D Max in MB</td><td class="cell"><%=RS0("HDD_Max")%></td>
        </tr>
        <tr>
        <td class="cellT">HD E used in MB</td><td class="cell"><%=RS0("HDE_Current")%></td>
        </tr>
        <tr>
        <td class="cellT">HD E Max in MB</td><td class="cell"><%=RS0("HDE_Max")%></td>
        </tr>
        <tr>
        <td class="cellT">S/N</td><td class="cell"><%=RS0("SN")%></td>
        </tr>
        <tr>
        <td class="cellT">Support Remotly</td><td class="cell">
        <%  IF RS0("Remote") = 1     Then %>
                <font color="green">YES</font>
        <%  ELSEIF RS0("Remote") = 0 Then %>
                <font color="Red">No</font>
        <%  Else %>
                UNKNOWN
        <%  End IF %>
        </td>
        </tr>
        <tr>
        <td class="cellT">Wight</td><td class="cell"><%=RS0("Wight")%></td>
        </tr>
        <tr>
        <td class="cellT">Rank</td><td class="cell"><%=RS0("Rank")%>%</td>
        </tr>
        <tr>
        <td class="cellT">Group</td><td class="cell"><%=RS0("Assets_Group")%></td>
        </tr>
        <tr>
        <td class="cellT">Is it same Owner?</td><td class="cell"></td>
        </tr>
        <tr>
        <td class="cellT">More Detailes</td><td class="cell"><a href="http://*********/default.aspx?item=quicksearch&q=<%=RS0("ISD_Name")%>"> Show </a></td> <!-- replace ***** with your tool's path -->
        </tr>
	</table>
   </td><td width="100">&nbsp;</td>
    <% End IF %>
   <td >

       <%IF RS1("Assets_Type_Id") = 1 Then %>
       <img src="img/pc.png"/>
       <%ElseIF RS1("Assets_Type_Id") = 2 Then %>
       <img src="img/laptop.png"  / >
       <%ElseIF RS1("Assets_Type_Id") = 3 Then %>
       <img src="img/monitor.png"/>
       <%ElseIF RS1("Assets_Type_Id") = 4 Then %>
       <img src="img/printer.png"/>
       <%ElseIF RS1("Assets_Type_Id") = 5 Then %>
       <img src="img/scanner.png"/>
       <%ElseIF RS1("Assets_Type_Id") = 6 Then %>
       <img src="img/iPad.png"/>
       <%Else %>
       <img src="img/no.png"/>
       <%End If %>
   </td></tr>
  </table>
        <% Else %>
        <h5> Error in finding asset's type</h5>
        <%End IF %>
        
        <%
        Set RS9 = Server.CreateObject("ADODB.RecordSet")
	    SQL9 = "SELECT * FROM  Assets where ISD_Name = '" & ISD & "'"
	    Set RS9 = con.execute(SQL9)
	    Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM  Owner where Assets_Id = '" & RS9("Assets_Id") & "' and Owner_Active='1'"
	    Set RS3 = con.execute(SQL3)
	    IF NOT RS3.EOF THEN
        Set RS4 = Server.CreateObject("ADODB.RecordSet")
    	SQL4 = "SELECT * FROM Emp e WHERE e.Emp_Id = '" & RS3("Emp_Id") & "'"
    	Set RS4 = con.execute(SQL4)
        %>
        <h2>Currently it is with </h2>
        <table>
        <tr> 
            <td><a href="Emp_Details.asp?Emp_Badge=<%=RS4("Emp_Badge")%>"><%=RS4("Emp_Name")%></a>&nbsp;</td>
            <td> Scince <%=RS3("Date_Reg")%></td>
            <td> , <%=RS3("Note")%></td>
            <td><input  type="submit" name="customers" value="<% = RS3("Owner_Id")%>" onclick="showCustomer(this.value)" /></td>
        </tr>
	</table>
    <div id="txtHint"> </div>


    <% If Emp_Id ="" Then %>
   <!-- <a href="Emp_Details.asp?Assets_Id=<%=RS9("Assets_Id")%>" target="_blank">Assign it for new Emp</a> -->
    <% Else %>
    <%
        Set RS44 = Server.CreateObject("ADODB.RecordSet")
    	SQL44 = "SELECT * FROM Emp e WHERE e.Emp_Id = '" & Emp_Id & "'"
    	Set RS44 = con.execute(SQL44)
        
    %>
    <form method="POST" action="Assign_Assets_O.asp" name="ISD" onsubmit="return show_alert()">
    <input type= "hidden" name = "Emp_Id" value="<% = RS3("Emp_Id") %>" >
    <input type= "hidden" name = "Assets_Id" value="<% = RS9("Assets_Id")%>" >
    <input type= "hidden" name = "Owner_Id" value="<% = RS3("Owner_Id")%>" >
    <input type= "hidden" name = "Badge" value="<% = RS44("Emp_Badge")%>" >
    <input type= "hidden" name = "Release_justification" value="<% response.write(Release_justification) %>" >
    <input type="submit" name="b" value="Confirm Assigning This Asset To <% response.write(RS44("Emp_Name")) %>"/>
    </form>
    <% End IF %>
    <h2>Asset's History</h2>
        <%
	    Set RS5 = Server.CreateObject("ADODB.RecordSet")
	    SQL5 = "SELECT * FROM  History where Assets_Id = '" & RS9("Assets_Id") & "' ORDER BY History_Id DESC"
	    Set RS5 = con.execute(SQL5)

	    IF NOT RS5.EOF THEN
        %>
        <table>
        <tr>
        <td class="TableTitle">Assigning Date</td>
        <td class="TableTitle">Owner</td>
        <td class="TableTitle">Note</td>
        </tr>
        <%Do While NOT RS5.EOF 
        Set RS6 = Server.CreateObject("ADODB.RecordSet")
    	SQL6 = "SELECT * FROM Emp e WHERE e.Emp_Id = '" & RS5("Emp_Id") & "' "
    	Set RS6 = con.execute(SQL6)     
        %>  
        <tr> 
            <td class="cell"><%=RS5("Date_Reg")%></td>
            <td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RS6("Emp_Badge")%>"><%=RS6("Emp_Badge")%></a>(<%=RS6("Emp_Name")%>)</td>
            <td class="cell"><%=RS5("Note")%></td>
        </tr>
        <%RS5.Movenext 
        loop
        %>
	</table>
    <%Else%>
        <h5>No History</h5>
    <%END IF %>
    <%Else%>
        <h5>Sorry , There is no recorde for this asset in the system , it must be regiserd firest by System Admin to confirme that</h5>
        <a href="Emp_Details.asp?Assets_Id=<%=RS9("Assets_Id")%>" target="_blank">Assign it for new Emp</a>
    <%END IF %>
    <br/><br/>
        <%Else
        RESPONSE.REDIRECT ("Asset_Details_Like.asp?ISD_Name=" & ISD & "&SType=" & SType & "")
    END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>