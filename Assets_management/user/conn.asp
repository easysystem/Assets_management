
<%

Dim Con
Dim Recordset
	ServerName	= "***" 'replace *** with your Server Name
	DBName		= "***" 'replace *** with your DB Name
	UserID		= "***" 'replace *** with your DB User name
	Password	= "***" 'replace *** with your DB User Password

'open the connection
Set Con = Server.CreateObject("ADODB.Connection")
Con.Open = "PROVIDER=SQLOLEDB;DATA SOURCE=" & ServerName & ";User Id=" & UserID & ";Password=" & Password & ";Initial Catalog=" & DBName & ";"

%>
