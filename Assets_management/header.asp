
<html dir="ltr"  >
<head>
<link rel="stylesheet" type="text/css" href="style.css"/>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1256"/>
 <script type="text/javascript">
     function show_alert() {
         if (document.Badge.Badge.value !== "") {

             var numericExpression = /^[0-9]+$/;
             if (document.Badge.Badge.value.match(numericExpression)) {
                 return true;
             } else {
                 alert("Sorry , Badge it must be numbers only");
                 Badge.Badge.focus();
                 return false;
             }
             
         } else
         return true;
     }

</script>
</head>
    <body >	 
    <table>
    <tr >
    <form method="POST" action="Emp_Details.asp" name="Badge" onsubmit="return show_alert()"  >
    <td>Badge #:</td>
    <td><input class="inputa" name="Badge" type="text" size="15" /></td>
    <td><input type="submit" name="b" value="Search BY Badge"/></td>
    </form>
    <form method="POST" action="Asset_Details.asp" name="ISD" >
    <td >ISD Name:</td>
    <td >
    <input class="inputa" name="ISD" type="text" size="15" />
    <input class="inputa" name="SType" type="hidden" value='1' />
    </td>
    <td><input type="submit" name="b" value="Search BY ISD"/></td>
    </form>
    <form method="POST" action="Emp_Details.asp" name="Name" >
    <td>Emp Name: </td>
    <td>
    <input class="inputa" name="Name" type="text" size="25" />
    <input class="inputa" name="SType" type="hidden" value='2' />
    </td>
    <td><input type="submit" name="b" value="Search BY Name"/></td>
    </form>
    <td><a href="index.asp">.: Home :.</a></td>
    <td><a href="index.asp">.: Reports :.</a></td>
    <td><a href="index.asp">.: statistics :.</a></td>
    <td><a href="logout.asp">.: LogOut :.</a></td>
    </tr>
    </table>
	</body>
</html>