 <!-- #include file="conn.asp" -->

 <%
 Emp_Badge        = REQUEST.FORM ("Emp_Badge")
 Dept_Id		  = REQUEST.FORM ("Dept_Id")
 Area_Id	      = REQUEST.FORM ("Area_Id")
 Emp_Title        = REQUEST.FORM ("Emp_Title")
 Emp_Name         = REQUEST.FORM ("Emp_Name")
 Emp_Ext          = REQUEST.FORM ("Emp_Ext")
 Emp_Pager        = REQUEST.FORM ("Emp_Pager")
 Emp_Office       = REQUEST.FORM ("Emp_Office")
 Emp_Active       = REQUEST.FORM ("Emp_Active")
 username         = REQUEST.FORM ("username")
 Note             = REQUEST.FORM ("Note")
 USER             = Session("SUSER_ID")


    'Check if the Employee if exist
 	Set RS0 = Server.CreateObject("ADODB.RecordSet")
	SQL0 = "SELECT * FROM  Emp where Emp_Badge = '" & Emp_Badge & "' or username = '" & username & "' "
	Set RS0 = con.execute(SQL0)
    If  RS0.EOF Then

        'Insert Data for new Emp
        SELECTION	= "SELECT * FROM Emp"
        SET RESULT	= con.EXECUTE (SELECTION)
        SET RESULT = Server.CreateObject("ADODB.RecordSet")
		RESULT.Open "Emp", con, 2, 3
	    RESULT.AddNew
			
			    RESULT  ("Emp_Name")      = Emp_Name
			    RESULT  ("Emp_Badge")     = Emp_Badge
                RESULT  ("Dept_Id")       = Dept_Id
                RESULT  ("Area_Id")       = Area_Id
                RESULT  ("Emp_Title")     = Emp_Title
                RESULT  ("Emp_Ext")	      = Emp_Ext
                RESULT  ("Emp_Pager")	  = Emp_Pager
                RESULT  ("Emp_Email")     = username+"@***" ' replace *** with Exchange server domain to get the full email
                RESULT  ("Emp_Office")    = Emp_Office
                RESULT  ("Emp_Active")	  = Emp_Active
                RESULT  ("username")	  = username
                RESULT  ("Emp_Reg")	      = now()
			    RESULT  ("Note")          = Note
			
	    RESULT.Update
        RESULT.Close

        'Check the Asset ID after inserting
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        SQL2 = "SELECT * FROM Emp where Emp_Badge = '" & Emp_Badge & "'"
        Set RS2 = con.execute(SQL2)
        Emp_Id = RS2("Emp_Id")


        'Add it to History
	    SELECTION	= "SELECT * FROM History"
	    SET RESULT	= con.EXECUTE (SELECTION)
		SET RESULT  = Server.CreateObject("ADODB.RecordSet")			
	    RESULT.Open "History", con, 2, 3
	    RESULT.AddNew

                RESULT  ("Emp_Id")	        = Emp_Id
                RESULT  ("Assets_Id")       = 0
			    RESULT  ("History_Active")	= 1
                RESULT  ("Date_Reg")	    = now()
			    RESULT  ("Note")            = "New Employee <a href=Emp_Details.asp?Emp_Badge=" & Emp_Badge & ">" + Emp_Name +"</a> ,  Inserted into System by USER ID ("& USER &")"
		
        RESULT.Update
		RESULT.Close
	Con.close
    RESPONSE.REDIRECT ("Emp_Add.asp?Note=Update_ok")
    Else %>
        <font color ="red">Pleas double check , The Employee is already Exist</font>
    <%END IF %>
