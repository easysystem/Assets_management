<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<%
	Content = ""							'Clear the Content string
	QStr = Request.QueryString("login")		'Save the login querystring to QStr
 %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Login to iSupport</title>
<link href="isupport1.css" rel="stylesheet" type="text/css" />
<script>
var is_chrome = navigator.userAgent.toLowerCase().indexOf('chrome');
if (!(is_chrome > 1)) {
alert('Use Google Chrome!')
}
</script>
</head>

<body>
<br /><br /><br /><br />
<table  align="center">
    <% if QStr <> "" Then %>
    <tr align="center">
        <td  align="center">
            <div id="login-box-error">
                <table >
				<tr >
				<td class="note" >
				<% if QStr="passfailed" then %>
				    Sorry Incorrect Pasword
				<%elseif QStr="namefailed" then%>
				    Sorry , username is not registerd
				<%elseif QStr="Time_Out" then%>
                    Sorry , you loged out tomutaically because either you are tring to access page not allowed for you or time out
				<%elseif QStr="Privilege" then%>
				    Sorry , you dont have acess for this page
				<%elseif QStr="NoData" then%>
				    Sorry you didnt login to iSupport Asset Manger
				<%elseif QStr="Bye" then%>
				    Bye
                <%end if%>
                </td>
				</tr>
                </table >
            </div>
        </td>
    </tr>
    <%Else%>
    <br /><br />
    <%end if%>
    <tr>
        <td>
        <br /><br /><br />
<div id="login-box">

<h2>Login</h2>
Use your company user name<br />

<form method="POST" action="login_operation.asp" name="login" onsubmit="return show_alert()" >
    <br />
        <div id="login-box-name" style="margin-top:20px;">user name:</div>
        <div id="login-box-field" style="margin-top:20px;"><input name="txtUsername" class="form-login" title="Username" value="" size="30" maxlength="2048" /></div>
        <div id="login-box-name" >Password:</div>
        <div id="login-box-field"> <input name="txtPassword" type="password" class="form-login" title="Password" value="" size="30" maxlength="2048" /></div>
    <br />
    <br />
    <br />
    <br />
        <input type="submit" class="loginButt" value="" name="b"/>
</form>
</div>
</td>
</tr>
</table>
</body>
</html>
