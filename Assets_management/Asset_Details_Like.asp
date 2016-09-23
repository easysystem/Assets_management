<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%ISD = Request.form("ISD")
IF ISD = "" Then
ISD = request.querystring("ISD_Name")
End If
%>
<%Note = request.querystring("Note")%>
<%SType = request.querystring("SType")%>
<%    
    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1 or  Session("SUser_Privilege") = 2 or  Session("SUser_Privilege") = 3  Then
%>
<%
if SType = 1 then
    vTitle = "Do you Mean : "
    vQuery = "SELECT * FROM  Assets where ISD_Name LIKE '%" & ISD & "%'"
Elseif SType = 2 then
    vTitle = "PCs and Laptops Need to fill its specifications :"
    vQuery = "SELECT * FROM  Assets where Assets_Processor='' and (Assets_Type_Id='1' or Assets_Type_Id = '2') and Assets_Active='1' ORDER  by Assets_Type_Id "
Elseif SType = 3 then
    vTitle = "PCs and Laptops LOST :"
    vQuery = "SELECT * FROM  Assets where (Assets_Active='2' or Assets_Active='3') ORDER  by Assets_Type_Id "
Elseif SType = 4 then
    vTitle = "PCs and Laptops DAMEGED :"
    vQuery = "SELECT * FROM  Assets where Assets_Active='4' ORDER  by Assets_Type_Id "
End IF

 %>
<html dir="ltr" >
    <head>
        <title>Asset Details</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>	 
		<br/><br/>
        <h1>Search</h1>
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = vQuery
	    Set RS0 = con.execute(SQL0)
	    IF NOT RS0.EOF THEN
        %>&nbsp;
        <h2><%Response.write(vTitle) %></h2>
    <table>
    <tr>
    <td class="TableTitle">#</td>
    <td class="TableTitle">ISD-Name</td>
    <td class="TableTitle">Type</td>
    <td class="TableTitle">Currently with</td>
    </tr>
     <%
     dim i 
     i=0
     Do While NOT RS0.EOF
     	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Assets_Type where Assets_Type_Id = '" & RS0("Assets_Type_Id") & "' "
	    Set RS1 = con.execute(SQL1)
     	Set RS2 = Server.CreateObject("ADODB.RecordSet")
	    SQL2 = "SELECT * FROM  Owner where Assets_Id = '" & RS0("Assets_Id") & "' and Owner_Active=1"
	    Set RS2 = con.execute(SQL2)

        
      %> 
        <tr>
        <td class="cell"><% i = i + 1 %><%response.write(i)%></td>
        <td class="cell"><a href="Asset_Details.asp?ISD=<%=RS0("ISD_Name")%>"><%=RS0("ISD_Name")%></a><a href="http://***********/default.aspx?item=quicksearch&q=<%=RS0("ISD_Name")%>"> ... </a></td> <!-- replace ***** with your tool's path -->
        <td class="cell"><%=RS1("Assets_Type_Name")%></td>
        <td class="cell">
        <%
        IF NOT RS2.EOF THEN
     	Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM  Emp where Emp_Id = '" & RS2("Emp_Id") & "' "
	    Set RS3 = con.execute(SQL3)
        %>
        <a href="Emp_Details.asp?Emp_Badge=<%=RS3("Emp_Badge")%>"><%=RS3("Emp_Name")%></a></td>
        <%Else%>
        <i>Not assigned</i>
        <%End If%>
        
        </tr>
        <%RS0.Movenext 
   loop
    %>	
	</table>
    <%Else%>
            <h5>Pleas double check of the ISD Name for the Asset</h5>
    <%END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>