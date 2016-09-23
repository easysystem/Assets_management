<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%
ISD = request.querystring("Emp_Id")
Note = request.querystring("Note")

    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1  Then
%>

<html dir="ltr" >
    <head>
        <title>Add New Asset</title>
        <link rel="stylesheet" type="text/css" href="style.css" />
        <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    <script type="text/javascript">
        function show_alert() {
            if (document.NewUser.Emp_Badge.value == null || document.NewUser.Emp_Badge.value == "")
            {
                alert("Badge Number Must be filled");
                NewUser.Emp_Badge.focus();
                return false;
            }
            if (document.NewUser.Emp_Badge.value !== "") {

                var numericExpression = /^[0-9]+$/;
                if (document.NewUser.Emp_Badge.value.match(numericExpression)) {
                    return true;
                } else {
                    alert("Sorry , Badge it must be numbers only");
                    NewUser.Emp_Badge.focus();
                    return false;
                }

            } else
                return true;
        }
    </script>
    </head>
    <body>	 
    <%	if Note = "Update_ok" then%><div class="note">Updated Successfully</div><% end if%>
		<br/><br/>
        <h1>New User</h1>
        <form method="POST" action="Emp_Add_O.asp" name="NewUser" onsubmit="return show_alert()" >
        <table>
        <tr>
        <td class="cellT">Name</td>
        <td class="cell"><input class="inputa" name="Emp_Name" type="text" size="40" /></td>
        </tr>
        <tr>
        <td class="cellT">Emp Badge</td>
        <td class="cell"><input class="inputa" name="Emp_Badge" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">User Name</td>
        <td class="cell"><input class="inputa" name="username" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">Title</td>
        <td class="cell">
            <select name="Emp_Title" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Emp_Title FROM Emp order by  Emp_Title "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Emp_Title")%>"><% = RS2("Emp_Title")%></option>
                <%RS2.Movenext
                loop%>
            </select>
        </td>
        </tr>
        <tr>
        <td class="cellT">Department</td>
        <td class="cell">
            <select name="Dept_Id" class="inputa">
                <%Set RS3 = Server.CreateObject("ADODB.RecordSet")
	            SQL3 = "SELECT * FROM Dept order by  Dept_Name "
	            Set RS3 = con.execute(SQL3)
                Do While NOT RS3.EOF%>	
                    <option value="<% = RS3("Dept_Id")%>"><% = RS3("Dept_Name")%></option>
                <%RS3.Movenext
                loop%>
            </select>
        </td>
        </tr>
        <tr>
        <td class="cellT">Area</td>
        <td class="cell">
            <select name="Area_Id" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT Area_Id FROM Area order by  Area_Id "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Area_Id")%>"><% = RS2("Area_Name")%></option>
                <%RS2.Movenext
                loop%>
            </select>
        </td>
        </tr>
        <tr>
        <td class="cellT">Office #</td>
        <td class="cell"><input class="inputa" name="Emp_Office" type="text" size="30" /></td>
        </tr>
        <tr>
        <td class="cellT">Ext.</td>
        <td class="cell"><input class="inputa" name="Emp_Ext" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">Pager</td>
        <td class="cell"><input class="inputa" name="Emp_Pager" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">Active</td>
        <td class="cell">
            <select name="Emp_Active" class="inputa">
    			<option value="1" selected>Yes</option>
    			<option value="0">No</option>
            </select>
         </td>
        </tr>
        <tr>
        <td class="cellT">Note</td>
        <td class="cell"><textarea name="Note" rows="5" cols="40" ></textarea></td>
        </tr>
	</table>
         <input name="SType" type="hidden" value='2' />
         <input type="submit" value="Insert" name="B3"/>
  </form>
 </body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>