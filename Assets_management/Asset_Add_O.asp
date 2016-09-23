 <!-- #include file="conn.asp" -->

 <%
 Assets_Type_Id        = REQUEST.FORM ("Assets_Type_Id")
 ISD_Name		       = REQUEST.FORM ("ISD_Name")
 Note       	       = REQUEST.FORM ("Note")
 Assets_Active	       = REQUEST.FORM ("Assets_Active")
 Assets_ModelF         = REQUEST.FORM ("Assets_ModelF")
 Assets_ModelL         = REQUEST.FORM ("Assets_ModelL")
 Assets_ProcessorF     = REQUEST.FORM ("Assets_ProcessorF")
 Assets_ProcessorL     = REQUEST.FORM ("Assets_ProcessorL")
 Assets_SpeedF         = REQUEST.FORM ("Assets_SpeedF")
 Assets_SpeedL         = REQUEST.FORM ("Assets_SpeedL")
 Ram_CurrentF          = REQUEST.FORM ("Ram_CurrentF")
 Ram_CurrentL          = REQUEST.FORM ("Ram_CurrentL")
 Ram_MaxF              = REQUEST.FORM ("Ram_MaxF")
 Ram_MaxL              = REQUEST.FORM ("Ram_MaxL")
 HDC_Current           = REQUEST.FORM ("HDC_Current")
 HDC_MAX               = REQUEST.FORM ("HDC_MAX")
 HDD_Current           = REQUEST.FORM ("HDD_Current")
 HDD_MAX               = REQUEST.FORM ("HDD_MAX")
 HDE_Current           = REQUEST.FORM ("HDE_Current")
 HDE_MAX               = REQUEST.FORM ("HDE_MAX")
 SN                    = REQUEST.FORM ("SN")
 Remote                = REQUEST.FORM ("Remote")
 USER                  = Session("SUSER_ID")

 IF Assets_Type_Id = 1 or Assets_Type_Id = 2 Then

     IF Assets_ModelF       = "" THEN Assets_Model = Assets_ModelL           ELSE Assets_Model = Assets_ModelF           END IF
     IF Assets_ProcessorF   = "" THEN Assets_Processor = Assets_ProcessorL   ELSE Assets_Processor = Assets_ProcessorF   END IF
     IF Assets_SpeedF       = "" THEN Assets_Speed = Assets_SpeedL           ELSE Assets_Speed = Assets_SpeedF           END IF
     IF Ram_CurrentF        = "" THEN Ram_Current = Ram_CurrentL             ELSE Ram_Current = Ram_CurrentF             END IF
     IF Ram_MaxF            = "" THEN Ram_Max = Ram_MaxL                     ELSE Ram_Max = Ram_MaxF                     END IF
     IF HDC_Current         = "" THEN HDC_Current   = 0 END IF
     IF HDC_MAX             = "" THEN HDC_MAX       = 0 END IF
     IF HDD_Current         = "" THEN HDD_Current   = 0 END IF
     IF HDD_MAX             = "" THEN HDD_MAX       = 0 END IF
     IF HDE_Current         = "" THEN HDE_Current   = 0 END IF
     IF HDE_MAX             = "" THEN HDE_MAX       = 0 END IF
 
 Else
 
     Assets_SpeedF       = 0
     Ram_CurrentF        = 0
     Ram_MaxF            = 0
     HDC_Current         = 0
     HDC_MAX             = 0
     HDD_Current         = 0
     HDD_MAX             = 0
     HDE_Current         = 0
     HDE_MAX             = 0

 END IF

    'Check if the ISD name exist
 	Set RS0 = Server.CreateObject("ADODB.RecordSet")
	SQL0 = "SELECT * FROM  Assets where ISD_Name = '" & ISD_Name & "'"
	Set RS0 = con.execute(SQL0)
    If  RS0.EOF Then


        'Insert Data for new Asset
        SELECTION	= "SELECT * FROM Assets"
        SET RESULT	= con.EXECUTE (SELECTION)
        SET RESULT = Server.CreateObject("ADODB.RecordSet")
		RESULT.Open "Assets", con, 2, 3
	    RESULT.AddNew
			
			    RESULT  ("Assets_Type_Id")	= Assets_Type_Id
                RESULT  ("ISD_Name")        = ISD_Name
                RESULT  ("Assets_Model")    = Assets_Model
                RESULT  ("Assets_Processor")= Assets_Processor
                RESULT  ("Assets_Speed")	= Assets_Speed
                RESULT  ("Ram_Current")	    = Ram_Current
                RESULT  ("Ram_Max")	        = Ram_Max
                RESULT  ("HDC_Current")	    = HDC_Current
                RESULT  ("HDC_MAX")	        = HDC_MAX
                RESULT  ("HDD_Current")	    = HDD_Current
                RESULT  ("HDD_MAX")	        = HDD_MAX
                RESULT  ("HDE_Current")	    = HDE_Current
                RESULT  ("HDE_MAX")	        = HDE_MAX
			    RESULT  ("Assets_Active")	= Assets_Active
                RESULT  ("Date_Reg")	    = now()
			    RESULT  ("Note")            = Note
                RESULT  ("SN")              = SN
                RESULT  ("Remote")          = Remote
			
	    RESULT.Update
        RESULT.Close

        'Check the Asset ID after inserting
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        SQL2 = "SELECT * FROM Assets where ISD_Name = '" & ISD_Name & "'"
        Set RS2 = con.execute(SQL2)
        Assets_Id = RS2("Assets_Id")


       'Assign it for inetial system

        SELECTION	= "SELECT * FROM Owner"
	    SET RESULT	= con.EXECUTE (SELECTION)
	    SET RESULT  = Server.CreateObject("ADODB.RecordSet")			
	    RESULT.Open "Owner", con, 2, 3
	    RESULT.AddNew
        
                RESULT  ("Assets_Id")	    = Assets_Id
                RESULT  ("Emp_Id")	        = 1
	            RESULT  ("Owner_Active")	= 1
                RESULT  ("Date_Reg")	    = now()
	            RESULT  ("Note")            = "Inserted into System BY USER ID ("& USER &")"
	            RESULT.Update

        'Add it to History
	    SELECTION	= "SELECT * FROM History"
	    SET RESULT	= con.EXECUTE (SELECTION)
		SET RESULT  = Server.CreateObject("ADODB.RecordSet")			
	    RESULT.Open "History", con, 2, 3
	    RESULT.AddNew

                RESULT  ("Assets_Id")	    = Assets_Id
			    RESULT  ("Emp_Id")	        = 1
			    RESULT  ("History_Active")	= 1
                RESULT  ("Date_Reg")	    = now()
			    RESULT  ("Note")            = "New Asset " + "<a href=Asset_Details.asp?ISD=" & ISD_Name & ">" + ISD_Name +"</a> , Inserted into System by USER ID ("& USER &")" 
		
        RESULT.Update
		RESULT.Close
	Con.close
    RESPONSE.REDIRECT ("Asset_Add.asp?Note=Update_ok")
    Else %>
        <font color ="red">Pleas double check , The ISD name is already Exist</font>
    <%END IF %>
