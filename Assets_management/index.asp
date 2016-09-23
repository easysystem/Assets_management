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

    'Assets's Function
    dim i,freex,freePC,freeLapTop,freeMoniter,freePrinter,freeScanner,freeiPad
    i = 1
    freePC=0
    freeLapTop=0
    freeMoniter=0
    freePrinter=0
    freeScanner=0
    freeiPad=0
    DO while i < 7
        freex= 0
    	Set RSf0 = Server.CreateObject("ADODB.RecordSet")
	    SQLf0 = "SELECT * FROM Assets WHERE Assets_Type_Id='"& (i) &"'"
	    Set RSf0 = con.execute(SQLf0)
        Do While NOT RSf0.EOF
            Set RS00 = Server.CreateObject("ADODB.RecordSet")
	        SQL00 = "SELECT * FROM Owner e WHERE Assets_Id = '" & RSf0("Assets_Id") & "' and Owner_Active = 1 and (e.Emp_Id = 2 or e.Emp_Id = 3  or e.Emp_Id = 4) "
	        Set RS00 = con.execute(SQL00)
            Do While NOT RS00.EOF
                freex = freex+1
            RS00.Movenext
            loop
        RSf0.Movenext
        loop
    If i = 1 then freePC = freex ELse If i = 2 Then freeLapTop = freex ELse If i = 3 Then freeMoniter = freex ELse If i = 4 Then freePrinter = freex ELse If i = 5 Then freeScanner = freex ELse If i = 6 Then freeiPad = freex End If
    i = i + 1
    loop
    	Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = "SELECT COUNT(Assets_Type_Id) as TotalPC FROM Assets WHERE Assets_Type_Id='1'"
	    Set RS0 = con.execute(SQL0)
    	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT COUNT(Assets_Type_Id) as TotalLapTop FROM Assets WHERE Assets_Type_Id='2'"
	    Set RS1 = con.execute(SQL1)
    	Set RS2 = Server.CreateObject("ADODB.RecordSet")
	    SQL2 = "SELECT COUNT(Assets_Type_Id) as TotalMoniter FROM Assets WHERE Assets_Type_Id='3'"
	    Set RS2 = con.execute(SQL2)     
    	Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT COUNT(Assets_Type_Id) as TotalPrinter FROM Assets WHERE Assets_Type_Id='4'"
	    Set RS3 = con.execute(SQL3)
    	Set RS4 = Server.CreateObject("ADODB.RecordSet")
	    SQL4 = "SELECT COUNT(Assets_Type_Id) as TotalScanner FROM Assets WHERE Assets_Type_Id='5'"
	    Set RS4 = con.execute(SQL4)
    	Set RS5 = Server.CreateObject("ADODB.RecordSet")
	    SQL5 = "SELECT COUNT(Assets_Type_Id) as TotaliPad FROM Assets WHERE Assets_Type_Id='6'"
	    Set RS5 = con.execute(SQL5)
    totalPC= RS0("TotalPC")
    totalLapTop= RS1("TotalLapTop")
    totalMoniter= RS2("TotalMoniter")
    totalPrinter= RS3("TotalPrinter")
    totalScanner= RS4("TotalScanner")
    totaliPad= RS5("TotaliPad")
    usedPC = totalPC - freePC
    usedLapTop = totalLapTop - freeLapTop
    usedMoniter = totalMoniter - freeMoniter
    usedPrinter = totalPrinter - freePrinter
    usedScanner = totalScanner - freeScanner
    usediPad = totaliPad - freeiPad
    perPC =FormatPercent(usedPC/totalPC)
    perPC1=FormatPercent(freePC/totalPC)
    perLapTop =FormatPercent(usedLapTop/totalLapTop)
    perLapTop1=FormatPercent(freeLapTop/totalLapTop)
    perMoniter =FormatPercent(usedMoniter/totalMoniter)
    perMoniter1=FormatPercent(freeMoniter/totalMoniter)
    perPrinter =FormatPercent(usedPrinter/totalPrinter)
    perPrinter1=FormatPercent(freePrinter/totalPrinter)
    perScanner =FormatPercent(usedScanner/totalScanner)
    perScanner1=FormatPercent(freeScanner/totalScanner)
    periPad =FormatPercent(usediPad/totaliPad)
    periPad1=FormatPercent(freeiPad/totaliPad)

%>
<html dir="ltr"  >
<head>
<link rel="stylesheet" type="text/css" href="style.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256"/>
</head>
    <body>	

    <br/><br/><br/><br/> 
    <table><tr><td>
    <table>
    <tr>
    <td><h2>Employee</h2></td>
    </tr>
    <tr >
    <td class="TableTitle">No PC</td>
    <td class="TableTitle">1 PC</td>
    <td class="TableTitle">> 1 PC</td>
    <td class="TableTitle">Total</td>
    </tr>
    <tr>
    <td class="Cell">62</td>
    <td class="Cell">155</td>
    <td class="Cell">23</td>
    <td class="Cell">230</td>
    </tr>
    </table>
    <br/>
    -- <a href="Asset_Details_Like.asp?SType=2" >PCs need to update</a>
    <br/><br/>
    -- <a href="AssetDept.asp" >Assets/Departments Distribution</a>
    <br/><br/>
    -- <a href="Assets_Spect.asp?SType=1" >Assets Specifications</a>
    <br/><br/>
    -- <a href="Asset_Details_Like.asp?SType=3" >Lost Assets</a>
    <br/><br/>
    -- <a href="Asset_Details_Like.asp?SType=4" >Damaged Assets</a>
    <br/><br/>
    -- <a href="Assets_Inv.asp?d=1" >Assets's Inventory</a>
    <br/><br/>
    -- <a href="Asset_Add.asp" >Add New Asset</a>
    <br/><br/>
    -- <a href="Emp_Add.asp" >Add New Employee</a>
    </td><td width='50'>
    </td><td>
       <h2>Last 10 History</h2>
        <%
	    Set RS55 = Server.CreateObject("ADODB.RecordSet")
	    SQL55 = "SELECT Top 10 * FROM  History where History_Active = '1' ORDER BY History_Id DESC"
	    Set RS55 = con.execute(SQL55)

	    IF NOT RS55.EOF THEN
        %>
        <table>
        <tr>
        <td class="TableTitle">Date</td>
        <td class="TableTitle">Owner</td>
        <td class="TableTitle">Note</td>
        </tr>
        <%Do While NOT RS55.EOF 
        Set RS66 = Server.CreateObject("ADODB.RecordSet")
    	SQL66 = "SELECT * FROM Emp e WHERE e.Emp_Id = '" & RS55("Emp_Id") & "' "
    	Set RS66 = con.execute(SQL66)     
        %>  
        <tr> 
            <td class="cell"><%=RS55("Date_Reg")%></td>
            <td class="cell"><a href="Emp_Details.asp?Emp_Badge=<%=RS66("Emp_Badge")%>"><%=RS66("Emp_Badge")%></a>(<%=RS66("Emp_Name")%>)</td>
            <td class="cell"><%=RS55("Note")%></td>
        </tr>
        <%RS55.Movenext 
        loop
        %>
	</table>
    <%Else%>
        <h5>No History</h5>
    <%END IF %>


    </td></tr>
    <tr><td>
    <table>
    <tr>
    <td><h2>Assets</h2></td>
    </tr>
    <tr>
    <td class="TableTitle">Type</td>
    <td class="TableTitle">Free</td>
    <td class="TableTitle">Assigned</td>
    <td class="TableTitle">Total</td>
    </tr>
    <tr >
    <td class="cellT">PC</td>
    <td class="cell"><%Response.Write(FreePC)%></td>
    <td class="cell"><%Response.Write(usedPC)%></td>
    <td class="cell"><%Response.Write(totalPC)%></td>
    </tr>
    <tr>
    <td class="cellT">LapTop</td>
    <td class="cell"><%Response.Write(FreeLapTop)%></td>
    <td class="cell"><%Response.Write(usedLapTop)%></td>
    <td class="cell"><%Response.Write(totalLapTop)%></td>
    </tr>
    <tr>
    <td class="cellT">Moniter</td>
    <td class="cell"><%Response.Write(FreeMoniter)%></td>
    <td class="cell"><%Response.Write(usedMoniter)%></td>
    <td class="cell"><%Response.Write(totalMoniter)%></td>
    </tr>
    <tr>
    <td class="cellT">Printer</td>
    <td class="cell"><%Response.Write(FreePrinter)%></td>
    <td class="cell"><%Response.Write(usedPrinter)%></td>
    <td class="cell"><%Response.Write(totalPrinter)%></td>
    </tr>
    <tr>
    <td class="cellT">Scanner</td>
    <td class="cell"><%Response.Write(FreeScanner)%></td>
    <td class="cell"><%Response.Write(usedScanner)%></td>
    <td class="cell"><%Response.Write(totalScanner)%></td>
    </tr>
    <tr>
    <td class="cellT">iPad</td>
    <td class="cell"><%Response.Write(FreeiPad)%></td>
    <td class="cell"><%Response.Write(usediPad)%></td>
    <td class="cell"><%Response.Write(totaliPad)%></td>
    </tr>
    </table>
    </td><td width='50'>
    </td><td>
        <table>
    <tr>
    <td><h2>Percentage</h2></td>
    </tr>
    <tr>
    <td class="TableTitle">Type</td>
    <td class="TableTitle">Assigned &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;
    &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp; Free</td>
    </tr>
    <tr >
    <td class="cellT">PC</td>
    <td class="cell"><%Response.Write(perPC)%>&emsp;
                <%ii = int (FormatNumber((usedPC/totalPC) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(perPC1)%></td>
    </tr>
    <tr>
    <td class="cellT">LapTop</td>
    <td class="cell"><%Response.Write(perLapTop)%>&emsp;
                <%ii = int (FormatNumber((usedLapTop/totalLapTop) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(perLapTop1)%></td>
    </tr>
    <tr>
    <td class="cellT">Moniter</td>
    <td class="cell"><%Response.Write(perMoniter)%>&emsp;
                <%ii = int (FormatNumber((usedMoniter/totalMoniter) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(perMoniter1)%></td>
    </tr>
    <tr>
    <td class="cellT">Printer</td>
    <td class="cell"><%Response.Write(perPrinter)%>&emsp;
                <%ii = int (FormatNumber((usedPrinter/totalPrinter) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(perPrinter1)%></td>
    </tr>
    <tr>
    <td class="cellT">Scanner</td>
    <td class="cell"><%Response.Write(perScanner)%>&emsp;
                <%ii = int (FormatNumber((usedScanner/totalScanner) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(perScanner1)%></td>
    </tr>
    <tr>
    <td class="cellT">iPad</td>
    <td class="cell"><%Response.Write(periPad)%>&emsp;
                <%ii = int (FormatNumber((usediPad/totaliPad) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                    %><font color ="red" ><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                    %><font color ="green" ><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write(periPad1)%></td>
    </tr>
    </table>
    
    </td></tr>
    </table>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>