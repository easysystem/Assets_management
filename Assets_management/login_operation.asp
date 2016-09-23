<!-- #include file="conn.asp" -->

<%
	ServerName	= "***" 'replace *** with your Server Name
	DBName		= "***" 'replace *** with your DB Name
	UserID		= "***" 'replace *** with your DB User name
	Password2	= "***" 'replace *** with your DB User Password
	sConnString = "PROVIDER=SQLOLEDB;DATA SOURCE=" & ServerName & ";UID=" & UserID & ";PWD=" & Password2 & ";DATABASE=" & DBName & ";"
	
	'Save the entered username and password
	Username = Request.Form("txtUsername")	
	Password = Request.Form("txtPassword")	

    set conn = Server.CreateObject("ADODB.Connection")
	conn.open sConnString
	Set RS = Server.CreateObject("ADODB.RecordSet")

	'Build connection with database
	rs.Open "SELECT * FROM SUser where SUser_Name='"& Username &"'", conn, 1 
	
	'If there is no record with the entered username, close connection
	'and go back to login with QueryString

	If rs.recordcount = 0 then
		rs.close
		conn.close
		set rs=nothing
		set conn=nothing
		Response.Redirect("login.asp?login=namefailed")
	end if

	'If it is removed user
	if rs("SUser_Active") = "0" then
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("login.asp?login=remove")
	end if

	'If entered password is right, close connection and open mainpage
	if  rs("SUser_Pass")          = Password then
		Session("SUser_name")     = rs("SUser_FullName")
		Session("SUser_ID")       = rs("SUser_ID")
		Session("SUser_Privilege")= rs("SUser_Privilege")
		Session("SUser_Active")   = rs("SUser_Active")
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("index.asp")
	'If entered password is wrong, close connection 
	'and return to login with QueryString
	else
		rs.Close
		conn.Close
		set rs=nothing
		set conn=nothing
		Response.Redirect("login.asp?login=passfailed")
	end if	
%>
