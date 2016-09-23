 <!-- #include file="conn.asp" -->

 <head>
 <script type="text/javascript">
     function show_alert() {
         if (document.Relase.Release_justification.value == "") {
             alert("Sorry , you didn't write the release justification");
             Relase.Release_justification.focus();
             return (false);
         }
         return true;
     }
</script>
 </head>

 <%
 response.write(Request.QueryString("q"))
 Set RS4 = Server.CreateObject("ADODB.RecordSet")
 SQL4 = "SELECT * FROM Owner e WHERE e.Owner_Id = '" & Request.QueryString("q") & "'"
 Set RS4 = con.execute(SQL4)
 %>
 <html>
 <body bgcolor="#fbc6cf">
   <b><font color="red">Release Justification</font></b>
   <form method="POST" action="Relase_O.asp" name="Relase" onsubmit="return show_alert()"  >
    <table bgcolor="#fec2d1">
     <tr><td>
      <textarea name="Release_justification" rows="5" cols="40" ></textarea>
      </td></tr>
      <tr><td>
      <input id="Radio1" type="radio" name="New_Emp_Id" value="111" />P/CABIN A
      <input id="Radio2" type="radio" name="New_Emp_Id" value="112" />P/CABIN B
      <input id="Radio3" type="radio" name="New_Emp_Id" value="113" />CPHHI
      <input id="Radio4" type="radio" name="New_Emp_Id" value="114" />Damaged
      <br/>
      or Badge #<input name="Badge" type="text" size="8" />
      </td></tr>
      <tr><td>
      <input type= "hidden" name = "Assets_Id" value="<% = RS4("Assets_Id")%>" >
      <input type= "hidden" name = "Emp_Id" value="<% = RS4("Emp_Id")%>" >
      <input type= "hidden" name = "Owner_Id" value="<% = RS4("Owner_Id")%>" >

      <input type="submit" name="b" value="submit"/>
      </td></tr>
    </table>
   </form>
        <%
	    Set RS0 = Server.CreateObject("ADODB.RecordSet")
	    SQL0 = "SELECT * FROM  Emp where Emp_ID = '" & Request.QueryString("q") & "'"
	    Set RS0 = con.execute(SQL0)

	    IF NOT RS0.EOF THEN
        %>&nbsp;
    <table>
        <tr>
        <td>Name:</td><td><%=RS0("Emp_Name")%></td>
        </tr>
        <tr>
        <td>Tittle</td><td><%=RS0("Emp_Title")%></td>
        </tr>
        <tr>
        <td>Badge</td><td><%=RS0("Emp_Badge")%></td>
        </tr>
        <tr>
        <td>Active</td><td><%=RS0("Emp_Active")%></td>
        </tr>
        <tr>
        <td>Note</td><td><%=RS0("Note")%></td>
        </tr>
	</table>
    <%Else%>
        <font color ="red">No Recode</font>
    <%END IF %>

    </body>
    </html>