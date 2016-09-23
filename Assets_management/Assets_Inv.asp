<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<%

DeptID = request.querystring("d")
    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1 or  Session("SUser_Privilege") = 2 or  Session("SUser_Privilege") = 3  Then

    Set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL1 = "SELECT * FROM Dept where dept_Id='" & DeptID & "' "
	Set RS1 = con.execute(SQL1)
%>

<html dir="ltr" >
    <head>
        <title>Assets</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>
    <br/><br/>
      <h1><%=RS1("Dept_Name")%>'s Assets</h1>
        <br/>
        <table>
            <tr>
                <td class="TableTitle">NO.</td>
                <td class="TableTitle">Badge</td>
                <td class="TableTitle">Employee Name</td>
                <td class="TableTitle">Office</td>
                <td class="TableTitle">Ext.</td>
                <td class="TableTitle">Asset Type</td>
                <td class="TableTitle">ISD-Name</td>
                <td class="TableTitle">Monitor-ISD</td>
                <td class="TableTitle" width="350">Note</td>
            </tr>
                <%
	            Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT * FROM Emp where dept_Id='" & DeptID & "' "
	            Set RS2 = con.execute(SQL2)
                dim i
                i = 0
                Do While NOT RS2.EOF

                	Set RS3 = Server.CreateObject("ADODB.RecordSet")
	                SQL3 = "SELECT * FROM Owner where Emp_Id='" & RS2("Emp_Id") & "' and Owner_Active=1"
	                Set RS3 = con.execute(SQL3)
                    IF NOT RS3.EOF THEN
                    Do While NOT RS3.EOF
                
                	    Set RS4 = Server.CreateObject("ADODB.RecordSet")
	                    SQL4 = "SELECT * FROM Assets where Assets_Id='" & RS3("Assets_Id") & "'"
	                    Set RS4 = con.execute(SQL4)
                        Do While NOT RS4.EOF
                            Set RS5 = Server.CreateObject("ADODB.RecordSet")
	                        SQL5 = "SELECT * FROM Assets_Type where Assets_Type_Id='" & RS4("Assets_Type_Id") & "'"
	                        Set RS5 = con.execute(SQL5)
                            %>
                            <tr>
                                <td class="cell" rowspan="2"><% i = i + 1 %><%response.write(i)%></td>
                                <td class="cell" rowspan="2"><%=RS2("Emp_Badge")%></td>
                                <td class="cell" rowspan="2"><a href="Emp_Details.asp?Emp_Badge=<%=RS2("Emp_Badge")%>"><%=RS2("Emp_Name")%></a></td>
                                <td class="cell"><%=RS2("Emp_Office")%></td>
                                <td class="cell"><%=RS2("Emp_Ext")%></td>
                                <td class="cell" rowspan="2"><%=RS5("Assets_Type_Name")%></td>
                                <td class="cell"><a href="Asset_Details.asp?ISD=<%=RS4("ISD_Name")%>"><%=RS4("ISD_Name")%></a></td>
                                <td class="cell" rowspan="2"></td>
                                <td class="cell" rowspan="2"></td>
                            </tr>
                            <tr>
                                <td class="cell"><br/><br/></td>
                                <td class="cell"></td>
                                <td class="cell"></td>
                            </tr>
                            <%
                        RS4.Movenext
                        loop
                    RS3.Movenext
                    loop
		    Else%>
                            <tr>
                                <td class="cell" rowspan="2"><% i = i + 1 %><%response.write(i)%></td>
                                <td class="cell" rowspan="2"><%=RS2("Emp_Badge")%></td>
                                <td class="cell" rowspan="2"><a href="Emp_Details.asp?Emp_Badge=<%=RS2("Emp_Badge")%>"><%=RS2("Emp_Name")%></a></td>
                                <td class="cell"><%=RS2("Emp_Office")%></td>
                                <td class="cell"><%=RS2("Emp_Ext")%></td>
                                <td class="cell" rowspan="2"><i><font color="red">No Asset</font></i></td>
                                <td class="cell" rowspan="2"></td>
                                <td class="cell" rowspan="2"></td>
                                <td class="cell" rowspan="2"></td>
                            </tr>
                            <tr>
                                <td class="cell"><br/><br/></td>
                                <td class="cell"></td>
                            </tr>
		   <%End IF
                RS2.Movenext
                loop
                %>
                <tr><td class="cell"><br/><br/></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td></tr>
                <tr><td class="cell"><br/><br/></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td></tr>
                <tr><td class="cell"><br/><br/></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td></tr>
                <tr><td class="cell"><br/><br/></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td></tr>                                                
                <tr><td class="cell"><br/><br/></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td><td class="cell"></td></tr>
            </table>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>