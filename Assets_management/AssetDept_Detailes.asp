<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%ISD = Request.form("ISD")
IF ISD = "" Then
ISD = request.querystring("ISD_Name")
End If
%>
<%AssetsType = request.querystring("AssetsType")%>
<%DeptID = request.querystring("DeptID")%>
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
     	Set RSa1 = Server.CreateObject("ADODB.RecordSet")
	    SQLa1 = "SELECT * FROM  Dept where Dept_Id = '" & DeptID & "' "
	    Set RSa1 = con.execute(SQLa1)
        Set RSa2 = Server.CreateObject("ADODB.RecordSet")
	    SQLa2 = "SELECT * FROM  Assets_Type where Assets_Type_Id = '" & AssetsType & "' "
	    Set RSa2 = con.execute(SQLa2)
        Set RSa3 = Server.CreateObject("ADODB.RecordSet")
	    SQLa3 = "SELECT * FROM  Emp where Emp_Badge = '" & RSa1("Emp_Id_Head") & "' and Emp_Active = 1  "
	    Set RSa3 = con.execute(SQLa3)

 %>
<html dir="ltr" >
    <head>
        <title>All <%=RSa2("Assets_Type_Name")%> Details in <%=RSa1("Dept_Name")%> Department</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>	 
		<br/><br/>
        <h1>Assets Details</h1>
        <%IF NOT RSa1.EOF THEN%>
        &nbsp;
        <h2><%=RSa2("Assets_Type_Name")%>s in <%=RSa1("Dept_Name")%> Department</h2>
    <table>
    <tr>
    <td class="TableTitle">ISD-Name</td>
    <td class="TableTitle">Badge</td>
    <td class="TableTitle">Currently with</td>
    <td class="TableTitle">Date</td>
    <td class="TableTitle">Signature</td>
    <td class="TableTitle">Note</td>
    </tr>

     <%
     Set RS0 = Server.CreateObject("ADODB.RecordSet")
	 SQL0 = "SELECT * FROM  Assets where Assets_Type_Id = '" & AssetsType & "' and Assets_Active=1"
	 Set RS0 = con.execute(SQL0)
     
     Do While NOT RS0.EOF

     	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Owner where Assets_Id = '" & RS0("Assets_Id") & "' and Owner_Active=1 ORDER BY Emp_Id"
	    Set RS1 = con.execute(SQL1)

     	Set RS2 = Server.CreateObject("ADODB.RecordSet")
	    SQL2 = "SELECT * FROM  Emp where Emp_Id = '" & RS1("Emp_Id") & "' and Emp_Active = 1  and dept_Id= '" & DeptID & "'  "
	    Set RS2 = con.execute(SQL2)
        
        IF not RS2.EOF THEN
      %> 
        <tr>
        <td class="cell"><a href="Asset_Details.asp?ISD=<%=RS0("ISD_Name")%>"><%=RS0("ISD_Name")%></a></td>
        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RS2("Emp_Badge")%>"><%=RS2("Emp_Badge")%></a></td>
        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RS2("Emp_Badge")%>"><%=RS2("Emp_Name")%></a></td>
        <td class="cell" width="70"></td>
        <td class="cell" width="90"></td>
        <td class="cell" width="260"></td>
        </tr>
        <%End if 
        RS0.Movenext 
        loop
        %>	
	</table>
    <br /><br /><br />

    <table>
    <tr>
    <td class="TableTitle">Department Manager</td>
    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RSa3("Emp_Badge")%>"><%=RSa3("Emp_Name")%></a></td>
    </tr>
    <tr>
    <td class="TableTitle">Date</td>
    <td class="cell"></td>
    </tr>
    <tr>
    <td class="TableTitle">Signuter</td>
    <td class="cell"></td>
    </tr>
    </table>
    <%Else%>
            <h5>The department that you are looking for is not exist in the system</h5>
    <%END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>