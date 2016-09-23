 <!-- #include file="conn.asp" -->

 <%
    
    group1 = 85             ' 15% For IT , Servers
    group2 = 75             ' 10% For Heads
    group3 = 10             ' 65% For Staff
    group4 = 5              ' 5% For Temp , the other 5% will be in store
    Assets_Speed_W = 20      ' Assets Speed Wight
    Ram_Current_W = 30       ' Ram Current Wight
    int numOFrec = 1         ' Number of assets recoeds
    int totalRank = 0        ' the total value of Asseys' rank
    int i = 0
    Rank_Type = request.querystring("Rank_Type")
	Set RS = Server.CreateObject("ADODB.RecordSet")
	SELECTION	= "SELECT * , (( '" & Assets_Speed_W & "' * Assets_Speed) + ( '" & Ram_Current_W & "' * Ram_Current) + HDC_Max + HDD_Max + HDE_Max) as aspectrank FROM  Assets where Assets_Type_Id = '" & Rank_Type & "' and  Assets_Active = '1' order by aspectrank desc "
	SET RESULT	= con.EXECUTE (SELECTION)
    Do While NOT RESULT.EOF   
        ar= RESULT("aspectrank")
        totalRank = totalRank + ar
        RESULT.Movenext
        numOFrec= numOFrec + 1
    loop
    avragespect = totalRank/numOFrec

	'SAVING TO THE DATABASE
	SET RESULT = Server.CreateObject("ADODB.RecordSet")
	RESULT.Open SELECTION, con, 2, 3
    Do While NOT RESULT.EOF             
        i = i + 1
        wight = RESULT("aspectrank")
        RESULT  ("Rank")= FormatNumber((wight/avragespect) * 100)
        RESULT  ("Wight")= wight
	    RESULT.Update
        RESULT.Movenext
    loop
	
	RESULT.Close


    Set RS2 = Server.CreateObject("ADODB.RecordSet")
	SELECTION2	= "SELECT * FROM  Assets where Assets_Type_Id = '" & Rank_Type & "' order by Wight asc "
	SET RESULT2	= con.EXECUTE (SELECTION2)
	'SAVING TO THE DATABASE
	SET RESULT2 = Server.CreateObject("ADODB.RecordSet")
	RESULT2.Open SELECTION2, con, 2, 3
        int ii = 0
    Do While NOT RESULT2.EOF
        wight_Color= RESULT2("Rank")             
        ii = ii + 1
        If ( 100*ii/numOFrec > group1 ) then 
            RESULT2  ("Rank")= "<font color =Green >"+ wight_Color +"</font>"
            RESULT2  ("Assets_Group")= "<font color =Green >Super</font>"
        elseIf ( 100*ii/numOFrec > group2 ) then
            RESULT2  ("Rank")= "<font color =Blue >"+ wight_Color +"</font>"
            RESULT2  ("Assets_Group")= "<font color =Blue >Very Good</font>"
        elseIf ( 100*ii/numOFrec > group3 ) then
            RESULT2  ("Rank")= "<font color =Black >"+ wight_Color +"</font>"
            RESULT2  ("Assets_Group")= "<font color =Black >Good</font>"
        elseIf ( 100*ii/numOFrec > group4 ) then
            RESULT2  ("Rank")= "<font color =SaddleBrown >"+ wight_Color +"</font>"
            RESULT2  ("Assets_Group")= "<font color =SaddleBrown >Need to Change</font>"
        else
            RESULT2  ("Rank")= "<font color =Red >"+wight_Color+"</font>"
            RESULT2  ("Assets_Group")= "<font color =Red >Temp</font>"
        End If
	    RESULT2.Update
        RESULT2.Movenext
    loop
    RESULT2.Close

	Con.close
	RESPONSE.REDIRECT ("Assets_Spect.asp?Note=Update_ok")
%>
