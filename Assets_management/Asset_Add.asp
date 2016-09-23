<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<!-- #include file="header.asp" -->
<%
ISD = request.querystring("ISD_Name")
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
            if (document.NewAsset.ISD_Name.value == null || document.NewAsset.ISD_Name.value == "")
            {
                alert("ISD Name Must be filled");
                NewAsset.ISD_Name.focus();
                return false;
            }
        }
    </script>
    </head>
    <body>	 
    <%	if Note = "Update_ok" then%><div class="note">Updated Successfully</div><% end if%>
		<br/><br/>
        <h1>New Asset</h1>
        <form method="POST" action="Asset_Add_O.asp" name="NewAsset" onsubmit="return show_alert()" >
        <table>
        <tr>
        <td class="cellT">Asset Type</td>
        <td class="cell">
            <select name="Assets_Type_Id" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT * FROM Assets_Type "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Assets_Type_Id")%>"><% = RS2("Assets_Type_Name")%></option>
                <%RS2.Movenext
                loop%>
            </select>
         </td>
        </tr>
        <tr>
        <td class="cellT">ISD - Name</td>
        <td class="cell"><input class="inputa" name="ISD_Name" type="text" size="15" /></td>
        </tr>
        <tr>
        <td class="cellT">Active</td>
        <td class="cell">
            <select name="Assets_Active" class="inputa">
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

		<br/>
        <h1>Asset Spect.</h1>
        <table>
        <tr>
        <td class="cellT">Model</td>
        <td class="cell">
            <select name="Assets_ModelL" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Assets_Model FROM Assets order by  Assets_Model "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Assets_Model")%>"><% = RS2("Assets_Model")%></option>
                <%RS2.Movenext
                loop%>
            </select>
             Other 
            <input class="inputa" name="Assets_ModelF" type="text" size="25" />
        </td>
        </tr>
        <tr>
        <td class="cellT">Processor Type</td>
        <td class="cell">
            <select name="Assets_ProcessorL" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Assets_Processor FROM Assets order by  Assets_Processor "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Assets_Processor")%>"><% = RS2("Assets_Processor")%></option>
                <%RS2.Movenext
                loop%>
            </select>
             Other 
            <input class="inputa" name="Assets_ProcessorF" type="text" size="25" />
        </td>
        </tr>
        <tr>
        <td class="cellT">Speed in MHz</td>
        <td class="cell">
            <select name="Assets_SpeedL" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Assets_Speed FROM Assets order by  Assets_Speed "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Assets_Speed")%>"><% = RS2("Assets_Speed")%></option>
                <%RS2.Movenext
                loop%>
            </select>
             Other 
            <input class="inputa" name="Assets_SpeedF" type="text" size="25" />
        </td>
        </tr>
        <tr>
        <td class="cellT">Current Ram MB</td>
        <td class="cell">
            <select name="Ram_CurrentL" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Ram_Current FROM Assets order by  Ram_Current "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Ram_Current")%>"><% = RS2("Ram_Current")%></option>
                <%RS2.Movenext
                loop%>
            </select>
             Other 
            <input class="inputa" name="Ram_CurrentF" type="text" size="25" />
        </td>
        </tr>
        <tr>
        <td class="cellT">Max Ram MB</td>
        <td class="cell">
            <select name="Ram_MaxL" class="inputa">
                <%Set RS2 = Server.CreateObject("ADODB.RecordSet")
	            SQL2 = "SELECT DISTINCT Ram_Max FROM Assets order by  Ram_Max "
	            Set RS2 = con.execute(SQL2)
                Do While NOT RS2.EOF%>	
                    <option value="<% = RS2("Ram_Max")%>"><% = RS2("Ram_Max")%></option>
                <%RS2.Movenext
                loop%>
            </select>
             Other 
            <input class="inputa" name="Ram_MaxF" type="text" size="25" />
        </td>
        </tr>
        <tr>
        <td class="cellT">HD C used in MB</td>
        <td class="cell"><input class="inputa" name="HDC_Current" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">HD C Max in MB</td>
        <td class="cell"><input class="inputa" name="HDC_MAX" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">HD D used in MB</td>
        <td class="cell"><input class="inputa" name="HDD_Current" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">HD D Max in MB</td>
        <td class="cell"><input class="inputa" name="HDD_MAX" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">HD E used in MB</td>
        <td class="cell"><input class="inputa" name="HDE_Current" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">HD E Max in MB</td>
        <td class="cell"><input class="inputa" name="HDE_MAX" type="text" size="10" /></td>
        </tr>
        <tr>
        <td class="cellT">Serial Number</td>
        <td class="cell"><input class="inputa" name="SN" type="text" size="30" /></td>
        </tr>
        <tr>
        <td class="cellT">Support Reomtly</td>
        <td class="cell">
            <select name="Remote" class="inputa">
                <option value="3">Don't Know</option>
               <option value="0">No</option>
               <option value="1">Yes</option>
            </select>
        </td>
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