<!-- #include file="conn.asp" -->

<%
dim Username
dim Password
dim Domain

	'Save the entered username and password
	Username = Request.Form("txtUsername")	
	Password = Request.Form("txtPassword")	
    Domain = "*****" ' replace *** with Exchange server domain example Jutail-dm.com.sa

    result = AuthenticateUser(UserName, Password, Domain)
    if result then
        response.write "<h3>Authentication Successfuly!</h3>" 
    else
        response.write "<h3>Authentication Failed!</h3>"    
    end if
    
function AuthenticateUser(UserName, Password, Domain)


        Set objConnection = server.CreateObject("ADODB.Connection")
        Set objCommand = CreateObject("ADODB.Command")
        objConnection.Provider = "ADsDSOObject"
        objConnection.Properties("Encrypt Password") = true
        objConnection.Open "DS Query", UserName, Password
        Set objCommand.ActiveConnection = objConnection
        objCommand.Properties("Page Size") = 1000
        ' replace *** correct path for LDAP and the domain name after the @
        strQuery =  "SELECT cn, mail, member, department ,extensionAttribute1, title  FROM 'LDAP://****/DC=***,DC=***,DC=***' " &_
                    "WHERE objectCategory='user' and MAIL='" & UserName & "@*****'  " &_
                    "ORDER BY cn"
        objCommand.CommandText =  strQuery

        On Error Resume Next
        Set objRecordSet = objCommand.Execute
        If Not objRecordSet.EOF Then 
            Response.Write (Err.Description& "<br><br>")
            If Err.Number <> 0 Then
                AuthenticateUser = false
                Response.Redirect("login2.asp?login=passfailed")
	        else%>
                <br/>Name: <%=objRecordSet.Fields("cn")%>
                <br/>Email: <%=objRecordSet.Fields("mail")%>
                <br/>Department: <%=objRecordSet.Fields("department")%>
                <br/>Badge: <%=objRecordSet.Fields("extensionAttribute1")%>
                <br/>Title: <%=objRecordSet.Fields("title")%>
            <%AuthenticateUser = True
            end if
        On Error GoTo 0
        Else
           Response.Redirect("login2.asp?login=namefailed")
        End If
       
end function
%>
