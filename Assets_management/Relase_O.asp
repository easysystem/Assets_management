 <!-- #include file="conn.asp" -->

 <%
 Assets_Id		       = REQUEST.FORM ("Assets_Id")
 Emp_Id		           = REQUEST.FORM ("Emp_Id")
 Owner_Id		       = REQUEST.FORM ("Owner_Id")
 New_Emp_Id            = REQUEST.FORM ("New_Emp_Id")
 Release_justification = REQUEST.FORM ("Release_justification")
 Badge                 = REQUEST.FORM ("Badge")

 If New_Emp_Id = "" Then
 New_Emp_Id = Badge
 End IF
 If Badge = "" Then
 Badge = New_Emp_Id
 End IF


    'get info for the new emp
 	Set RS0 = Server.CreateObject("ADODB.RecordSet")
	SQL0 = "SELECT * FROM  Emp where Emp_Badge = '" & Badge & "'"
	Set RS0 = con.execute(SQL0)
    If Not RS0.EOF THen
    New_Name = RS0("Emp_Name")
    New_Badge = RS0("Emp_Badge")
    New_Id = RS0("Emp_Id")

    'get info for the old emp
 	Set RS1 = Server.CreateObject("ADODB.RecordSet")
	SQL1 = "SELECT * FROM  Emp where Emp_Id = '" & Emp_Id & "'"
	Set RS1 = con.execute(SQL1)
    Old_Name = RS1("Emp_Name")
    Old_Badge = RS1("Emp_Badge")

	SELECTION	= "SELECT * FROM Owner e WHERE e.Owner_Id = '" & Owner_Id & "'"
	SET RESULT	= con.EXECUTE (SELECTION)

	'SAVING TO THE DATABASE
	SET RESULT = Server.CreateObject("ADODB.RecordSet")
	RESULT.Open SELECTION, con, 2, 3

			RESULT  ("Owner_Active")	= 0

	RESULT.Update
	RESULT.Close

    
    Set RS = Server.CreateObject("ADODB.RecordSet")	

	SELECTION	= "SELECT * FROM History"
	SET RESULT	= con.EXECUTE (SELECTION)
			SET RESULT = Server.CreateObject("ADODB.RecordSet")			
			RESULT.Open "History", con, 2, 3
			RESULT.AddNew
			

			RESULT  ("Emp_Id")	        = Emp_Id
			RESULT  ("Assets_Id")	    = Assets_Id
			RESULT  ("Date_Reg")	    = now()
			RESULT  ("History_Active")	= 1
			RESULT  ("Note")       = "Assets relased because of <b>" + Release_justification + "</b> And assigned to " + "<a href=Emp_Details.asp?Emp_Badge=" & New_Badge & ">"+ New_Name +"</a>"
			
			RESULT.Update
			RESULT.Close


   ' new emp assign

    SELECTION	= "SELECT * FROM Owner"
	SET RESULT	= con.EXECUTE (SELECTION)
	SET RESULT = Server.CreateObject("ADODB.RecordSet")			
	RESULT.Open "Owner", con, 2, 3
	RESULT.AddNew


    RESULT  ("Emp_Id")	        = New_Id
    RESULT  ("Assets_Id")	    = Assets_Id
	RESULT  ("Owner_Active")	= 1
    RESULT  ("Date_Reg")	    = now()
	RESULT  ("Note")            = "Asset assigned because of <b>" + Release_justification + "</b> And it was with " + "<a href=Emp_Details.asp?Emp_Badge=" & Old_Badge & ">"+ Old_Name +"</a>"
	RESULT.Update

  
    Set RS = Server.CreateObject("ADODB.RecordSet")	
	SELECTION	= "SELECT * FROM History"
	SET RESULT	= con.EXECUTE (SELECTION)
			SET RESULT = Server.CreateObject("ADODB.RecordSet")			
			RESULT.Open "History", con, 2, 3
			RESULT.AddNew
			

			RESULT  ("Emp_Id")	        = New_Id
			RESULT  ("Assets_Id")	    = Assets_Id
			RESULT  ("Date_Reg")	    = now()
			RESULT  ("History_Active")	= 1
			RESULT  ("Note")       = "Asset assigned because of <b>" + Release_justification + "</b> And it was with " + "<a href=Emp_Details.asp?Emp_Badge=" & Old_Badge & ">"+ Old_Name +"</a>"
			
			RESULT.Update
			RESULT.Close

 If New_Emp_Id = "114" Then

	SELECTION	= "SELECT * FROM Assets e WHERE e.Assets_Id = '" & Assets_Id & "'"
	SET RESULT	= con.EXECUTE (SELECTION)

	'SAVING TO THE DATABASE
	SET RESULT = Server.CreateObject("ADODB.RecordSet")
	RESULT.Open SELECTION, con, 2, 3
			RESULT  ("Assets_Active")	= 4

	RESULT.Update
	RESULT.Close


 End IF

     
     
     ' Send notification =======================================================================================

    Dim objEmail , strHTML, Emp_name , Emp_email, asset_Type , ISD_Name , ISD_NameN , reg_Date
    Dim strHTML2a ,strHTML2b, strHTML2c, strHTML2d, strHTML2e ' The body of email
    strHTML2a = " "
    strHTML2b = " "
    strHTML2c = " "
    strHTML2d = " "
    strHTML2e = " "
    strBodyMax= " "

     
        Set RS111 = Server.CreateObject("ADODB.RecordSet") ' to get the isd name for relased 
	    SQL111 = "SELECT * FROM Assets e WHERE e.Assets_Id = '" & Assets_Id & "'"
	    Set RS111 = con.execute(SQL111)
        isd_name_relased  = RS111("ISD_Name")


        Set objEmail = CreateObject("CDO.Message")
        Set RS11 = Server.CreateObject("ADODB.RecordSet") ' to get the emp detailes
	    SQL11 = "SELECT * FROM  Emp where Emp_Id = '" & Emp_Id & "'"
	    Set RS11 = con.execute(SQL11)
        Emp_name = RS11("Emp_Name")
        Emp_email =RS11("Emp_email")
        Emp_ide  =RS11("Emp_id")

        strHead = "<html lang=""en""><head><meta http-equiv=""content-type"" content=""text/html; charset=iso-8859-1"" /></head>" & chr(13) &_
        "<title>Your record has been updated in iSupport</title>" & chr(13) &_
        "<body style=""background-color: #F3F3F3 ; font-family: 'Segoe UI Semilight', 'Open Sans', Verdana, Arial, Helvetica, sans-serif; font-weight: 500; font-size: 11pt; letter-spacing: 0.02em"">" & chr(13) &_
        "<br>Dear <b><i>"& Emp_name &" </i></b><br><br><br>New update has been done in my records."& chr(13) &_
        "<br>I just remove the flowing asset from under your name " & chr(13) &_
        " <b><i><font style=""color: #4477bb;""> "& isd_name_relased &"</font> </i></b>.<br><br>Now you have the following  assets :"

	    Set RS1 = Server.CreateObject("ADODB.RecordSet")
	    SQL1 = "SELECT * FROM  Owner where Emp_ID = '" & Emp_ide & "' and Owner_Active='1' ORDER by Owner_Id DESC" ' to get all active assets
	    Set RS1 = con.execute(SQL1)
	    IF NOT RS1.EOF THEN
            Set RS22 = Server.CreateObject("ADODB.RecordSet") ' to get the last asset
    	    SQL22 = "SELECT * FROM Assets e WHERE e.Assets_Id = '" & RS1("Assets_Id") & "'"
    	    Set RS22 = con.execute(SQL22)
            ISD_NameN = RS22("ISD_Name")
  

           strbody1 = " <table><tbody><tr style=""font-family: 'Segoe UI Semilight', 'Open Sans', Verdana, Arial, Helvetica, sans-serif; font-weight: 500; font-size: 11pt; letter-spacing: 0.02em; line-height: 20px; padding: 3px 10px; border-right: 1px #ddd solid; border-bottom: 1px #ddd solid; box-sizing: border-box; background-color: #4477bb;  color:#ffffff;"" ><th >Type</th><th >ISD-Name</th><th >Assigning Date</th></tr>" 
            Do While NOT RS1.EOF 
                Set RS2 = Server.CreateObject("ADODB.RecordSet")
    	        SQL2 = "SELECT * FROM Assets e WHERE e.Assets_Id = '" & RS1("Assets_Id") & "'"
    	        Set RS2 = con.execute(SQL2)
                Set RS3 = Server.CreateObject("ADODB.RecordSet")
	            SQL3 = "SELECT * FROM Assets_Type e WHERE e.Assets_Type_Id = '" & RS2("Assets_Type_Id") & "'"
	            Set RS3 = con.execute(SQL3)
                asset_Type = RS3("Assets_Type_Name")
                ISD_Name = RS2("ISD_Name")
                reg_Date = RS1("Date_Reg")

                strHTML2 = "<tr style=""font-family: 'Segoe UI Semilight', 'Open Sans', Verdana, Arial, Helvetica, sans-serif; font-weight: 500; font-size: 11pt; letter-spacing: 0.02em; line-height: 20px; padding: 3px 10px;background-color: #fff"">" & chr(13) &_
                            "<td >" & asset_Type & "</td><td >" & ISD_Name & "</td><td >" & reg_Date & "</td></tr>" 

                            'Array to keep the number of assets can show in the email
                            iF      strHTML2a = " " Then
                              strHTML2a = strHTML2
                            Elseif  strHTML2b = " " Then
                              strHTML2b = strHTML2
                            Elseif  strHTML2c = " " Then
                              strHTML2c = strHTML2 
                            Elseif  strHTML2d = " " Then
                              strHTML2d = strHTML2
                            Elseif  strHTML2e = " " Then
                              strHTML2e = strHTML2
                              strBodyMax = "<br> You have special case, some of your assets not apper in this email , you can contact us to know what are they<br>"
                            End IF

            RS1.Movenext 
            loop
        strbody10=  "</tbody></table>"
        strBody = strbody1 + strHTML2a + strHTML2b + strHTML2c + strHTML2d + strHTML2e+ strbody10+ strBodyMax

    Else
        strBody = "So , Now you don't have any IT Asset ."
    END IF 
    
    strFoot = "<font style=""font-weight: 500; font-size: 9pt""><br> Note: This is an automated mail ,if you think this mail came to you by mistake <br>Don’t hesitate to call 40640" & chr(13) &_
              "<br><br><hr><table><tr><td><img style=""margin-left : 10px ; margin-right:10px"" src=""http://10.3.0.172/iSupport/img/Blogo.png""/></td><td>iSupport<br>Thanks</td></tr>"& chr(13) &_
              "<tr style=""font-weight: 500; font-size: 8pt""><td colspan=""2""> under evaluation copy</td></tr></table></body></html>"
    strHTML = strHead + strBody +strFoot

objEmail.From = "*******" ' replace ***** with sender's email example aaa@b.com
objEmail.To = emp_email
objEmail.CC = "********" ' replace ***** with sender's email example aaa@b.com
objEmail.Subject = "iSupport: Your Record Has Been Updated"
objEmail.HTMLBody  = cstr(strHTML)

With objEmail.Configuration.Fields
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
           .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "***"   ' replace *** with Exchange server IP
            .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "***" ' replace *** with sender user name only
            .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "***" ' replace *** with the password for the user
             .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 2 'it was 1 , but if it is 2 it does work
            .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
            .Update
End With
objEmail.Send
Set objEmail = Nothing
     '===============================================================
			Con.close

	RESPONSE.REDIRECT ("Emp_Details.asp?Emp_Badge=" & Badge & "&Note=Update_ok")

    Else %>
        <font color ="red">Pleas double check of the badge for the new employee</font>
    <%END IF %>
