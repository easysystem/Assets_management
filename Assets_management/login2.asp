<% @ LANGUAGE=VBScript CodePage="1256" %>
<!-- #include file="conn.asp" -->
<%
	BackgroundColor="#AFD1F8"
	BorderColor="#000080"
	
	Content = ""							'Clear the Content string
	QStr = Request.QueryString("login")		'Save the login querystring to QStr

 %>
<html>
<head>
<link rel="stylesheet" type="text/css" href="style.css">
<meta http-equiv="Content-Type"  content="text/html; charset=windows-1256">
<title>Login to Asset Managmer</title>
    <style type="text/css">
        .style1
        {
            height: 29px;
        }
        .style2
        {
            height: 18px;
        }
    </style>
</head>
<body>

<br><br><br><br>

				<% if QStr="passfailed" then %>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Sorry Incorrect Pasword
				</td>
				</tr>
</table>
				
				<%elseif QStr="namefailed" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Sorry , username is not registerd
				</td>
				</tr>
</table>
				
				<%elseif QStr="remove" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Sorry , your account is suspended
				</td>
				</tr>
</table>
				<%elseif QStr="Time_Out" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
                Sorry , you loged out tomutaically because either you are tring to access page not allowed for you or time out
				</td>
				</tr>
</table>
				<%elseif QStr="Privilege" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Sorry , you dont have acess for this page
				</td>
				</tr>
</table>
				<%elseif QStr="NoData" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Sorry you didnt login to iSupport Asset Manger
				</td>
				</tr>
</table>
				<%elseif QStr="Bye" then%>
<table align="center"  width="500" >
				<tr >
				<td class="note" >
				Bye
				</td>
				</tr>
</table>
				<% end if%>



<br><br>
			<table align="center" cellspacing="0" cellpadding="0" border="0"  style="font-size: x-small">
			

			<tr>
				<td  background="img/R.jpg" >&nbsp;</td>
				<td  bgcolor="#a8e6ff" >
				
				
			<b>	
			<p class="texta" >&nbsp;<br/><br/>
			User Name:
			<br>
            <form method="POST" action="login_operation2.asp" name="login" onsubmit="return show_alert()" >
			<input  type="text" name="txtUsername" size="25" class="inputa">
			<br>Password:
			<br><input type="password" name="txtPassword" size="25" class="inputa">
			<br><br><input type="submit" value="Log In" name="b"  style="float: right; border: 1px solid #C0C0C0; background-color: #F9F9F9; font-family: Arial, Helvetica, sans-serif;">
			</form><br><br>
			</b></font>
</p>			
				</td>
				<td  background="img/L.jpg" >&nbsp;</td>
			</tr>

		    
            </table>
</body>
</html>
