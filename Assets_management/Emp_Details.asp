<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%Badge = Request.form("Badge")
IF Badge = "" Then
Badge = request.querystring("Emp_Badge")
End If
Name = Request.Form("Name")
SType = Request.Form("SType")
%>
<%Note = request.querystring("Note")%>
<%    
    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1 or  Session("SUser_Privilege") = 2 or  Session("SUser_Privilege") = 3  Then
%>
<html dir="ltr" >
    <head>
        <title>Emp Details</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
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
<script type="text/javascript">
    function NewAssets(ida) {
        if (ida == "") {
            document.getElementById("txtHint2").innerHTML = "";
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
                document.getElementById("txtHint2").innerHTML = xmlhttp.responseText;
            }
        }
        xmlhttp.open("GET", "New_Assets.asp?q=" + ida, true);
        xmlhttp.send();
    }
</script>
    </head>
    <body>	 
        <%	if Note = "Update_ok" then%><div class="note">Updated Successfully</div><% end if%>
		<br/><br/>
        <h1>Employee information</h1>
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = "SELECT * FROM  Emp where Emp_Badge = '" & Badge & "'"
	    Set RS0 = con.execute(SQL0)

	    IF NOT RS0.EOF THEN

	    Set RS11 = Server.CreateObject("ADODB.RecordSet")
	    SQL11 = "SELECT * FROM  Dept where Dept_ID = '" & RS0("Dept_ID") & "'"
	    Set RS11 = con.execute(SQL11)

	    Set RS12 = Server.CreateObject("ADODB.RecordSet")
	    SQL12 = "SELECT * FROM  Emp where Emp_Badge = '" & RS11("Emp_ID_Head") & "'"
	    Set RS12 = con.execute(SQL12)
        	    
        %>&nbsp;
    <table>
        <tr>
        <td class="cellT">Name:</td><td class="cell"><%=RS0("Emp_Name")%></td>
        </tr>
        <tr>
        <td class="cellT">Tittle</td><td class="cell"><%=RS0("Emp_Title")%></td>
        </tr>
        <tr>
        <td class="cellT">Badge</td><td class="cell"><%=RS0("Emp_Badge")%></td>
        </tr>
        <tr>
        <td class="cellT">Email</td><td class="cell"><%=RS0("Emp_Email")%></td>
        </tr>
        <tr>
        <td class="cellT">Department</td><td class="cell"><% IF NOT RS11.EOF THEN %><%=RS11("Dept_Name")%><%Else %> Unknown <% End IF%></td>
        </tr>
        <tr>
        <td class="cellT">Area</td><td class="cell"><%=RS0("Area_Id")%></td>
        </tr>
        <tr>
        <td class="cellT">Office#</td><td class="cell"><%=RS0("Emp_Office")%></td>
        </tr>
        <tr>
        <td class="cellT">Ext.</td><td class="cell"><%=RS0("Emp_Ext")%></td>
        </tr>
        <tr>
        <td class="cellT">Pager</td><td class="cell"><%=RS0("Emp_Pager")%></td>
        </tr>
        <tr>
        <td class="cellT">Boss</td><td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RS12("Emp_Badge")%>"><%=RS12("Emp_Name")%></a></td>
        </tr>
        <tr>
        <td class="cellT">Registered</td><td class="cell"><%=RS0("Emp_Reg")%></td>
        </tr>
        <tr>
        <td class="cellT">Active</td><td class="cell"><%=RS0("Emp_Active")%></td>
        </tr>
        <tr>
        <td class="cellT">Note</td><td class="cell"><%=RS0("Note")%></td>
        </tr>
        <tr>
        <td class="cellT">Assets used</td><td class="cell">
        <a href="http://*******/default.aspx?item=userdetail&username=<%=RS0("username")%>&userdomain=*****">Show</a></td> <!-- replace ***** with your tool's path and domain in exchange -->
        </tr>
	</table>

        <h2> His/Her Assets</h2>
        <%
	    Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Owner where Emp_ID = '" & RS0("Emp_ID") & "' and Owner_Active='1'"
	    Set RS1 = con.execute(SQL1)
	    IF NOT RS1.EOF THEN
        %>
        <table>
        <tr>
        <td class="TableTitle">Type</td>
        <td class="TableTitle">ISD-Name</td>
        <td class="TableTitle">Assigning</td>
        <td class="TableTitle">Note</td>
        <td class="TableTitle">Relase</td>
        </tr>
        <%Do While NOT RS1.EOF 
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
    	SQL2 = "SELECT * FROM Assets e WHERE e.Assets_Id = '" & RS1("Assets_Id") & "'"
    	Set RS2 = con.execute(SQL2)
        Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM Assets_Type e WHERE e.Assets_Type_Id = '" & RS2("Assets_Type_Id") & "'"
	    Set RS3 = con.execute(SQL3)
        %>  
        <tr> 
            <td class="cell"><%=RS3("Assets_Type_Name")%></td>
            <td class="cell"><a href="Asset_Details.asp?ISD=<%=RS2("ISD_Name")%>"><%=RS2("ISD_Name")%></td>
            <td class="cell"><%=RS1("Date_Reg")%></td>
            <td class="cell"><%=RS1("Note")%></td>
            <td class="cell"><input  type="submit" name="customers" value="<% = RS1("Owner_Id")%>" onclick="showCustomer(this.value)" /></td>
        </tr>
        <%RS1.Movenext 
        loop
        %>
	</table>
    <div id="txtHint"> </div>
    <%Else%>
        <h5>No Assets</h5>
    <%END IF %>
    <br/>
Assign New Assets for
<input  type="submit" name="customers" value="<% = RS0("Emp_Id")%>" onclick="NewAssets(this.value)" />
<div id="txtHint2"> </div>
    <h2>His/Her History</h2>
        <%
	    Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM  History where Emp_ID = '" & RS0("Emp_ID") & "' ORDER BY History_Id DESC"
	    Set RS3 = con.execute(SQL3)

	    IF NOT RS3.EOF THEN
        %>        
        <table>
        <tr>
        <td class="TableTitle">Type</td>
        <td class="TableTitle">ISD-Name</td>
        <td class="TableTitle">Assigning</td>
        <td class="TableTitle">Note</td>
        </tr>
        <%Do While NOT RS3.EOF 
        aaa = RS3("Assets_Id")
        IF aaa = 0 Then
        assetHistory = "<i>No Asset</i>"
        assettype = "<i>No Asset</i>"
        Else
        Set RS4 = Server.CreateObject("ADODB.RecordSet")
    	SQL4 = "SELECT * FROM Assets e WHERE e.Assets_Id = '" & RS3("Assets_Id") & "'"
    	Set RS4 = con.execute(SQL4)
        Set RS7 = Server.CreateObject("ADODB.RecordSet")
    	SQL7 = "SELECT * FROM Assets_Type e WHERE e.Assets_Type_Id = '" & RS4("Assets_Type_Id") & "' "
    	Set RS7 = con.execute(SQL7)
        assetHistory = "<a href=Asset_Details.asp?ISD=" & RS4("ISD_Name") & ">" & RS4("ISD_Name") & "</a>"
        assettype = RS7("Assets_Type_Name")
        End IF
        %>  
        <tr> 
            <td class="cell"><%response.Write(assettype) %></td>
            <td class="cell"><%response.Write(assetHistory) %></td>
            <td class="cell"><%=RS3("Date_Reg")%></td>
            <td class="cell"><%=RS3("Note")%></td>
        </tr>
        <%RS3.Movenext 
        loop
        %>
	</table>
    <%Else%>
        <h5>No History</h5>
    <%END IF %>
    <br/><br/>
    <%Else
        If SType = 2 then
        RESPONSE.REDIRECT ("Emp_Details_Like.asp?Emp_Name=" & Name & "&SType=" & SType & "")
        else
        RESPONSE.REDIRECT ("Emp_Details_Like.asp?Emp_Badge=" & Badge & "")
        End If
    END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>