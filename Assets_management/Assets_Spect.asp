<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%ISD = Request.form("ISD")
IF ISD = "" Then
ISD = request.querystring("ISD_Name")
End If
%>
<%SType = Request.querystring("SType")
IF SType = "" Then
SType = request.form("SType")
End If
%>
<%Note = request.querystring("Note")%>
<%OrderBy = request.form("OrderBy")%>
<%OrderType = request.form("OrderType")%>
<%AssetType = request.form("Assets_Type_Id")%>
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
        <title>Asset Spect</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>
    <%	if Note = "Update_ok" then%><div class="note">Updated Successfully</div><% end if%>
		<br/><br/>
        <h1>Asset Spect</h1>
        <form method="POST" action="Assets_Spect.asp" name="Report" onsubmit="return show_alert()" >
        <select name="Assets_Type_Id" class="inputa">
                <option value="">All Assets</option>
                <%
	            Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT * FROM Assets_Type "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF
                %>	
                <option value="<% = RS2("Assets_Type_Id")%>"><% = RS2("Assets_Type_Name")%></option>
                <%
                RS2.Movenext
                loop
                %>
         </select>
         <select name="OrderBy" class="inputa">
                <option value="Wight">Wight</option>
                <option value="ISD_Name">ISD Name</option>
                <option value="Last_Name">Last User use</option>
                <option value="Assets_Model">Model</option>
                <option value="Assets_Processor">Processor Type</option>
                <option value="Assets_Speed">Speed</option>
                <option value="Ram_Current">Ram</option>
                <option value="">Hard Disk</option>
         </select>
         <select name="OrderType" class="inputa">
    			<option value="ASC">Ascending</option>
    			<option value="DESC">Descending</option>
         </select>
         <input name="SType" type="hidden" value='2' />
         <input type="submit" value="Generate" name="B3">
         <a href='Rank_O.asp?Rank_Type=1'> Rerank PCs </a>
          .||. 
         <a href='Rank_O.asp?Rank_Type=2'> Rerank Laptops </a>
         </form>
         
        <%
    if SType = 1 then
     %>
            <h5>Shows to generate the report</h5>
   <%
Elseif SType = 2 then
    vTitle = "Do you Mean : "
    vQuery = "SELECT * FROM  Assets where Assets_Type_Id LIKE '%" & AssetType & "%' ORDER BY " & OrderBy & " " & OrderType & ""

	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = vQuery
	    Set RS0 = con.execute(SQL0)
	    IF NOT RS0.EOF THEN
        %>&nbsp;
        <h2>Here is a report for <%Response.write(AssetType) %> arranged by <%Response.write(OrderBy) %> <%Response.write(vTitle) %></h2>
    <table>
    <tr>
    <td class="TableTitle">NO.</td>
    <td class="TableTitle">ISD-Name</td>
    <td class="TableTitle">Rank</td>
    <td class="TableTitle">Wight</td>
    <td class="TableTitle">Last User Use</td>
    <td class="TableTitle">Model</td>
    <td class="TableTitle">Processor</td>
    <td class="TableTitle">Speed</td>
    <td class="TableTitle">RAM</td>
    <td class="TableTitle">Hard Disk</td>
    <td class="TableTitle">HD Used</td>
    <td class="TableTitle">SN</td>
    <td class="TableTitle">Registration Date</td>
    <td class="TableTitle">Currently with</td>
    <td class="TableTitle">Note</td>
    </tr>
     <%
     dim i
        i = 0
     Do While NOT RS0.EOF
     	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Owner where Assets_Id = '" & RS0("Assets_Id") & "' and Owner_Active=1"
	    Set RS1 = con.execute(SQL1)
        IF  RS0("HDC_Current") > 1 Then HDCC = RS0("HDC_Current") Else HDCC = 0 End IF
        IF  RS0("HDD_Current") > 1 Then HDCD = RS0("HDD_Current") Else HDCD = 0 End IF
        IF  RS0("HDE_Current") > 1 Then HDCE = RS0("HDE_Current") Else HDCE = 0 End IF
        IF  RS0("HDC_Max") > 1 Then HDMC = RS0("HDC_Max") Else HDMC = 0 End IF
        IF  RS0("HDD_Max") > 1 Then HDMD = RS0("HDD_Max") Else HDMD = 0 End IF
        IF  RS0("HDE_Max") > 1 Then HDME = RS0("HDE_Max") Else HDME = 0 End IF
        HDC = ((HDCC+HDCD+HDCE)/1024)
        HDM = ((HDMC+HDMD+HDME+1)/1024)
        HDCG= round(HDC,2)
        HDMG= round(HDM,2)
        HDP = FormatPercent(HDC/HDM)
      %> 
        <tr>
        <td class="cell"><% i = i + 1 %><%response.write(i)%></td>
        <td class="cell"><a href="Asset_Details.asp?ISD=<%=RS0("ISD_Name")%>"><%=RS0("ISD_Name")%></a></td>
        <td class="cell"><%=RS0("Rank")%>%</td>
        <td class="cell"><%=RS0("Wight")%></td>
        <td class="cell"><%=RS0("Last_Name")%></td>
        <td class="cell"><%=RS0("Assets_Model")%></td>
        <td class="cell"><%=RS0("Assets_Processor")%></td>
        <td class="cell"><%=RS0("Assets_Speed")%></td>
        <td class="cell"><%=(RS0("Ram_Current")/1024)%> GB</td>
        <td class="cell"><%response.write(HDCG)%> GB</td>
        <td class="cell"><%response.write(HDP)%></td>
        <td class="cell"><%=RS0("SN")%></td>
        <td class="cell"><%=RS0("Date_Reg")%></td>
        <td class="cell">
        <%
        IF NOT RS1.EOF THEN
     	Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM  Emp where Emp_Id = '" & RS1("Emp_Id") & "' "
	    Set RS3 = con.execute(SQL3)
        %>
        <a href="Emp_Details.asp?Emp_Badge=<%=RS3("Emp_Badge")%>"><%=RS3("Emp_Name")%></a></td>
        <%Else%>
        <i>Not assigned</i>
        <%End If%>
        <td class="cell"><%=RS0("Note")%></td>
        </tr>
        <%RS0.Movenext 
   loop
    %>	
	</table>
    <%Else%>
            <h5>Pleas double check of the ISD Name for the Asset</h5>
    <%END IF %>
     <%END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>