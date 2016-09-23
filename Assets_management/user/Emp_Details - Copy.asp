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
        <title>Your Details</title>
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
        ' Check the user to give him his info. or somebody else
        If (Badge = Session("SUser_ID") )Then 
            IsSame= True
            Else IsSame= False
        End IF
        %>&nbsp;
        <br/>
        <h1>
        <%If IsSame Then %>
            Your
         <%Else %>
            His/Her
        <%End If %>
        information</h1>
<table><tr><td>
    <table class="striped2">
        <tr>
        <td>Name:</td><td><%=RS0("Emp_Name")%></td>
        </tr>
        <tr>
        <td >Tittle</td><td><%=RS0("Emp_Title")%></td>
        </tr>
        <tr>
        <td >Badge</td><td ><%=RS0("Emp_Badge")%></td>
        </tr>
        <tr>
        <td >Email</td><td ><%=RS0("Emp_Email")%></td>
        </tr>
        <tr>
        <td >Department</td><td ><% IF NOT RS11.EOF THEN %><%=RS11("Dept_Name")%><%Else %> Unknown <% End IF%></td>
        </tr>
        <tr>
        <td >Area</td><td ><%=RS0("Area_Id")%></td>
        </tr>
        <tr>
        <td >Office#</td><td ><%=RS0("Emp_Office")%></td>
        </tr>
        <tr>
        <td >Ext.</td><td ><%=RS0("Emp_Ext")%></td>
        </tr>
        <tr>
        <td >Pager</td><td ><%=RS0("Emp_Pager")%></td>
        </tr>
        <tr>
        <td>Boss</td><td><a href="Emp_Details.asp?Emp_Badge=<%=RS12("Emp_Badge")%>"><%=RS12("Emp_Name")%></a></td>
        </tr>
        <tr>
        <td>Registered</td><td ><%=RS0("Emp_Reg")%></td>
        </tr>
        <tr>
        <td >Active</td><td ><%=RS0("Emp_Active")%></td>
        </tr>
        <tr>
        <td >Note</td><td ><%=RS0("Note")%></td>
        </tr>
	</table>
</td><td width="350px">
    <div class="tile bg-color-red">
        <div class="tile-content">
            If you planning to arrange meeting and you need IT tech support or IT asset ,.. etc clik here
        </div>
        <div class="brand">
            <div class="name">Support Meeting</div>
        </div>
    </div>
    <div class="tile bg-color-green">
        <div class="tile-content">
            If you would like to request new IT asset such as PC, Laptop, Keyboard ,.. etc click here
        </div>
        <div class="brand">
            <div class="name">New Asset</div>
        </div>
    </div>
                            <div class="tile double bg-color-blue" data-role="tile-slider" data-param-period="3000">
                                    <div class="tile-content">
                                        <h2>iSupport@*****.com</h2>
                                        <h5>Leave a feedback</h5>
                                        <p>
                                           If you want to evaluate my services or suggest any thing , click here
                                        </p>
                                    </div>
                                    <div class="tile-content">
                                        <h2>iSupport@*****.com</h2>
                                        <h5>Leave a feedback</h5>
                                        <p>
                                           If you want to evaluate my services or suggest any thing , click here
                                        </p>
                                    </div>
                                    <div class="brand">
                                        <div class="badge">12</div>
                                        <img class="icon" src="images/home.png"/>
                                    </div>
                                </div>
</td></tr></table>
        <h2>Your Assets</h2>
        <%
	    Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Owner where Emp_ID = '" & RS0("Emp_ID") & "' and Owner_Active='1'"
	    Set RS1 = con.execute(SQL1)
	    IF NOT RS1.EOF THEN
        %>
        <table class="striped">
            <tbody>
        <tr>
        <th>Type</th>
        <th>ISD-Name</th>
        <th>Assigning</th>
        <th>Action</th>
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
            <td ><%=RS3("Assets_Type_Name")%></td>
            <td ><a href="Asset_Details.asp?ISD=<%=RS2("ISD_Name")%>"><%=RS2("ISD_Name")%></td>
            <td ><%=RS1("Date_Reg")%></td>
            <td><label class="input-control switch">
        <input type="checkbox">
        <span class="helper">Confirm you have this asset to start supporting you</span>
    </label></td>

        </tr>
        <%RS1.Movenext 
        loop
        %>
            </tbody>
	</table>
    <div id="txtHint"> </div>
    <%Else%>
        <h5>No Assets</h5>
    <%END IF %>
    <br/>
Request New IT Asset
<input  type="submit" name="customers" value="<% = RS0("Emp_Id")%>" onclick="NewAssets(this.value)" />
<div id="txtHint2"> </div>
    <h2>Your History</h2>
        <%
	    Set RS3 = Server.CreateObject("ADODB.RecordSet")
	    SQL3 = "SELECT * FROM  History where Emp_ID = '" & RS0("Emp_ID") & "' ORDER BY History_Id DESC"
	    Set RS3 = con.execute(SQL3)

	    IF NOT RS3.EOF THEN
        %>
        <table class="striped">
            
        <tr >
        <th >Type</th>
        <th >ISD-Name</th>
        <th >Assigning Date</th>
        <th >Note</th>
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
            <td><%response.Write(assettype) %></td>
            <td><%response.Write(assetHistory) %></td>
            <td><%=RS3("Date_Reg")%></td>
            <td><%=RS3("Note")%></td>
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