<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<%    
    'Log In	
    IF Session("SUser_ID") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=NoData")
    End If
    IF Session("SUser_Privilege") = "" Then
    RESPONSE.REDIRECT ("Login.asp?login=Time_Out")
    ElseIF  Session("SUser_Privilege") = 1 or  Session("SUser_Privilege") = 2 or  Session("SUser_Privilege") = 3  Then
%>
<html dir="ltr"  >
<head>
<link rel="stylesheet" type="text/css" href="stylescsb.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256"/>
</head>
    <body>	 
    <table>
    <tr>
    <form method="POST" action="Emp_Details.asp" name="Badge" onsubmit="return show_alert()"  >
    <td>Badge #:</td>
    <td><input class="inputa" name="Badge" type="text" size="25" /></td>
    <td><input type="submit" name="b" value="Search BY Badge"/></td>
    </form>
    </tr>
    <tr>
    <form method="POST" action="Asset_Details.asp" name="ISD" onsubmit="return show_alert()"  >
    <td>ISD Name:</td>
    <td><input class="inputa" name="ISD" type="text" size="25" /></td>
    <td><input type="submit" name="b" value="Search BY ISD"/></td>
    </form>
    </tr>
    </table>
	</body>
</html>
<!-- #include file="footer.asp" -->
<%
Else
RESPONSE.REDIRECT ("Login.asp?login=Privilege")
End IF
%>