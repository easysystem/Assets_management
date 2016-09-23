<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<%Emp_Id = request.querystring("Emp_Id")%>
<% If Emp_Id="" Then %>
<%Emp_Id = Request.QueryString("q")%>
<%End If %>
<html dir="ltr"  >
<head>
<link rel="stylesheet" type="text/css" href="stylescsb.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256"/>
</head>
    <body>
    <form method="POST" action="Asset_Details.asp" name="ISD" onsubmit="return show_alert()"  >
    <table>
    <tr>
    <td>ISD Name:</td>
    <td><input class="inputa" name="ISD" type="text" size="25" /></td>
    </tr>
    <tr>
    <td colspan="2">
    <textarea name="Release_justification" rows="5" cols="40" ></textarea>
    </td>
    </tr>
    <tr>
    <td colspan="2">
        <input type= "hidden" name = "Emp_Id" value="<% response.write(Emp_Id) %>" >
        <input type="submit" name="b" value="Assign"/>
    </td>
    </tr>
    </form>
    </tr>
    </table>
	</body>
</html>