<%@  language="VBScript" codepage="1256" %>

<html dir="ltr">
<head>
    <link rel="stylesheet" type="text/css" href="style.css" />
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
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
<body>

    <table>
        <tr>
            <form method="POST" action="Emp_Details.asp" name="Badge" onsubmit="return show_alert()">
                <td>Badge #:</td>
                <td>
                    <input class="inputa" name="Badge" type="text" size="15" /></td>
                <td>
                    <input type="submit" name="b" value="Search BY Badge" /></td>
            </form>
            <form method="POST" action="Asset_Details.asp" name="ISD">
                <td>ISD Name:</td>
                <td>
                    <input class="inputa" name="ISD" type="text" size="15" />
                    <input class="inputa" name="SType" type="hidden" value='1' />
                </td>
                <td>
                    <input type="submit" name="b" value="Search BY ISD" /></td>
            </form>
            <form method="POST" action="Emp_Details.asp" name="Name">
                <td>Emp Name: </td>
                <td>
                    <input class="inputa" name="Name" type="text" size="25" />
                    <input class="inputa" name="SType" type="hidden" value='2' />
                </td>
                <td>
                    <input type="submit" name="b" value="Search BY Name" /></td>
            </form>
            <td><a href="index.asp">.: Home :.</a></td>
            <td><a href="index.asp">.: Reports :.</a></td>
            <td><a href="index.asp">.: statistics :.</a></td>
            <td><a href="logout.asp">.: LogOut :.</a></td>
        </tr>
    </table>
    <p>
            &nbsp;</p>
		<br/><br/>
        <h1>Asset information</h1>
  <table >
   <tr><td>
    <table>
        <tr>
        <td class="cellT">ISD Name:</td><td class="cell">Rl - 10321</td>
        </tr>
        <tr>
        <td class="cellT">Type</td><td class="cell">Laptop</td>
        </tr>
        <tr>
        <td class="cellT">Assets Active</td><td class="cell">
                <font color="green">YES</font>
        </td>
        </tr>
        <tr>
        <td class="cellT">Registration Date</td><td class="cell">11-02-2010</td>
        </tr>
        <tr>
        <td class="cellT">Note</td><td class="cell">has HD port</td>
        </tr>
	</table>
   </td>
   <td width="100">&nbsp;</td>
   <td valign=bottom>
        <table valign=bottom>
        <tr>
        <td class="cellT">Last Login By</td><td class="cell">alhazmilo</td>
        </tr>
        <tr>
        <td class="cellT">Model</td><td class="cell">HP - 1010C</td>
        </tr>
        <tr>
        <td class="cellT">Processor Type</td><td class="cell">i7</td>
        </tr>
        <tr>
        <td class="cellT">Speed in GHz</td><td class="cell">2.3</td>
        </tr>
        <tr>
        <td class="cellT">Current Ram</td><td class="cell">8</td>
        </tr>
        <tr>
        <td class="cellT">Max Ram</td><td class="cell">16</td>
        </tr>
        <tr>
        <td class="cellT">HD C used in MB</td><td class="cell">11520</td>
        </tr>
        <tr>
        <td class="cellT">HD C Max in MB</td><td class="cell">32645</td>
        </tr>
        <tr>
        <td class="cellT">HD D used in MB</td><td class="cell">320</td>
        </tr>
        <tr>
        <td class="cellT">HD D Max in MB</td><td class="cell">5000</td>
        </tr>
        <tr>
        <td class="cellT">HD E used in MB</td><td class="cell">1320</td>
        </tr>
        <tr>
        <td class="cellT">HD E Max in MB</td><td class="cell">10000</td>
        </tr>
        <tr>
        <td class="cellT">S/N</td><td class="cell">21345421357132</td>
        </tr>
        <tr>
        <td class="cellT">Support Remotly</td><td class="cell">
                <font color="green">YES</font>
        </td>
        </tr>
        <tr>
        <td class="cellT">Wight</td><td class="cell">1.2</td>
        </tr>
        <tr>
        <td class="cellT">Rank</td><td class="cell">83 %</td>
        </tr>
        <tr>
        <td class="cellT">Group</td><td class="cell">A</td>
        </tr>
        <tr>
        <td class="cellT">Is it same Owner?</td><td class="cell"></td>
        </tr>
        <tr>
        <td class="cellT">More Detailes</td><td class="cell"><a href="http://*tem=quicksearch&q=3"> Show </a></td>
        </tr>
	</table>
   </td><td width="100">&nbsp;</td>
   <td >

       <img src="img/laptop.png"  / >

   </td></tr>
  </table>
       

        <h2>Currently it is with </h2>
        <table>
        <tr> 
            <td><a href="Emp_Details.asp?Emp_Badge=1">12412</a>&nbsp;</td>
            <td> Scince 20-06-2013</td>
            <td> , as he requested &nbsp;&nbsp;</td> 
            <td> <input  type="submit" name="customers" value="Release" onclick="showCustomer(this.value)" /></td>
        </tr>
	</table>
    <div id="txtHint"> </div>


    <form method="POST" action="Assign_Assets_O.asp" name="ISD" onsubmit="return show_alert()">

    </form>
    <h2>Asset's History</h2>

        <table>
        <tr>
        <td class="TableTitle">Assigning Date</td>
        <td class="TableTitle">Owner</td>
        <td class="TableTitle">Note</td>
        </tr>

        <tr> 
            <td class="cell">20-10-2013 13:12:13</td>
            <td class="cell"><a href="Emp_Details.asp?Emp_Badge=1">1234 </a> Ali Hafez</td>
            <td class="cell">as he requested</td>
        </tr>
        <tr> 
            <td class="cell">15-08-2013 10:17:39</td>
            <td class="cell"><a href="Emp_Details.asp?Emp_Badge=1">6512 </a> Mohammed Al Salman</td>
            <td class="cell">Project closed</td>
        </tr>
        <tr> 
            <td class="cell">08-09-2011 09:52:14</td>
            <td class="cell"><a href="Emp_Details.asp?Emp_Badge=1">1234 </a> Saleh Al Onazi</td>
            <td class="cell">HD size not enough</td>
        </tr>

	</table>

	</body>
</html>
