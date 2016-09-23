<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%Badge = Request.form("Badge")
IF Badge = "" Then
Badge = request.querystring("Emp_Badge")
End If
%>
<%Note = request.querystring("Note")%>
<%SType = request.querystring("SType")%>
<%Name = request.querystring("Emp_Name")%>
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
        <title>Result for Searching about an Employee </title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>	 
		<br/><br/>
        <h1>Search</h1>
        
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
        If Stype = 2 then 
        SQL0 = "SELECT * FROM  Emp where Emp_Name LIKE '%" & Name & "%'"
        Else
	    SQL0 = "SELECT * FROM  Emp where Emp_Badge LIKE '%" & Badge & "%'"
        End If
	    Set RS0 = con.execute(SQL0)

	    IF NOT RS0.EOF THEN
        %>&nbsp;
        <h2>Do You Mean : </h2>
    <table class="striped">
    <tr>
    <th>#</th>
    <th>Badge</th>
    <th>Name</th>
    </tr>

     <%
     dim i 
    i=0     
     Do While NOT RS0.EOF %> 
        <tr>
        <td><% i = i + 1 %><%response.write(i)%></td>
        <td><a href="Emp_Details.asp?Emp_Badge=<%=RS0("Emp_Badge")%>"><%=RS0("Emp_Badge")%></a></td>
        <td><%=RS0("Emp_Name")%></td>
        </tr>
        <%RS0.Movenext 
   loop
    %>	
	</table>
    <%Else%>
            <h5>Pleas double check of the badge for the employee</h5>
    <%END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>