<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->

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
        <title>Assets / Departments</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    </head>
    <body>	 
		<br/><br/>
        <h1>Assets / Departments</h1>
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = "SELECT * FROM  Dept where Dept_Active = '1'"
	    Set RS0 = con.execute(SQL0)
	    IF NOT RS0.EOF THEN
        %>&nbsp;
        <h2>Assets / Departments Distribution</h2>
    <table>
    <tr>
    <td class="TableTitle">Department-Name</td>
    <td class="TableTitle">PCs</td>
    <td class="TableTitle">Laptops</td>
    <td class="TableTitle">Moniter</td>
    <td class="TableTitle">Printer</td>
    <td class="TableTitle">Scanner</td>
    <td class="TableTitle">iPad</td>
    <td class="TableTitle">Total</td>
    </tr>
     <%
     Do While NOT RS0.EOF
     	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Emp where Emp_Active =1 and Dept_Id = '" & RS0("Dept_Id") & "' "
	    Set RS1 = con.execute(SQL1)
        countPC = 0
        countLaptop = 0
        countMoniter = 0
        countPrinter = 0
        countScanner = 0
        countiPad = 0
        totalAssetDept = 0
        Do While NOT RS1.EOF
     	    Set RS2 = Server.CreateObject("ADODB.RecordSet")
	        SQL2 = "SELECT * FROM  Owner where Emp_Id = '" & RS1("Emp_Id") & "' and Owner_Active=1"
	        Set RS2 = con.execute(SQL2)
            
            Do While NOT RS2.EOF
                Set RS3 = Server.CreateObject("ADODB.RecordSet")
	            SQL3 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=1 and Assets_Active=1"
	            Set RS3 = con.execute(SQL3)
                Do While NOT RS3.EOF
                    RS3.Movenext
                    countPC = countPC + 1
                loop
                Set RS4 = Server.CreateObject("ADODB.RecordSet")
	            SQL4 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=2 and Assets_Active=1"
	            Set RS4 = con.execute(SQL4)
                Do While NOT RS4.EOF
                    RS4.Movenext
                    countLaptop = countLaptop + 1
                loop
                Set RS5 = Server.CreateObject("ADODB.RecordSet")
	            SQL5 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=3 and Assets_Active=1"
	            Set RS5 = con.execute(SQL5)
                Do While NOT RS5.EOF
                    RS5.Movenext
                    countMoniter = countMoniter + 1
                loop
                Set RS6 = Server.CreateObject("ADODB.RecordSet")
	            SQL6 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=4 and Assets_Active=1"
	            Set RS6 = con.execute(SQL6)
                Do While NOT RS6.EOF
                    RS6.Movenext
                    countPrinter = countPrinter + 1
                loop
                Set RS7 = Server.CreateObject("ADODB.RecordSet")
	            SQL7 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=5 and Assets_Active=1"
	            Set RS7 = con.execute(SQL7)
                Do While NOT RS7.EOF
                    RS7.Movenext
                    countScanner = countScanner + 1
                loop
                Set RS8 = Server.CreateObject("ADODB.RecordSet")
	            SQL8 = "SELECT * FROM  Assets where Assets_Id = '" & RS2("Assets_Id") & "' and Assets_Type_Id=6 and Assets_Active=1"
	            Set RS8 = con.execute(SQL8)
                Do While NOT RS8.EOF
                    RS8.Movenext
                    countiPad = countiPad + 1
                loop
                totalAssetDept=countPC+countLaptop+countMoniter+countPrinter+countScanner+countiPad
            RS2.Movenext
            loop
        RS1.Movenext
        loop

        
      %> 
        <tr>
        <td class="cell"><%=RS0("Dept_Name")%></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=1&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(CountPC)%></a></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=2&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(countLaptop)%></a></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=3&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(countMoniter)%></a></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=4&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(countPrinter)%></a></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=5&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(countScanner)%></a></td>
        <td class="cell"><a href="AssetDept_Detailes.asp?AssetsType=6&DeptID=<%=RS0("Dept_ID")%>"><%Response.Write(countiPad)%></a></td>
        <td class="cell"><a href="Assets_Inv.asp?d=<%=RS0("Dept_ID")%>"><%Response.Write(totalAssetDept)%></a></td>
        </tr>
        <%RS0.Movenext 
    loop
    %>	
	</table>
    <%Else%>
            <h5>No Department in the system</h5>
    <%END IF %>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>