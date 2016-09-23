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


    <br />
    <br />
    <br />
    <br />
    <table>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>
                            <h2>Employee</h2>
                        </td>
                    </tr>
                    <tr>
                        <td class="TableTitle">No PC</td>
                        <td class="TableTitle">1 PC</td>
                        <td class="TableTitle">> 1 PC</td>
                        <td class="TableTitle">Total</td>
                    </tr>
                    <tr>
                        <td class="Cell">62</td>
                        <td class="Cell">155</td>
                        <td class="Cell">23</td>
                        <td class="Cell">230</td>
                    </tr>
                </table>
                <br />
                -- <a href="Asset_Details_Like.asp?SType=2">PCs need to update</a>
                <br />
                <br />
                -- <a href="AssetDept.asp">Assets/Departments Distribution</a>
                <br />
                <br />
                -- <a href="Assets_Spect.asp?SType=1">Assets Specifications</a>
                <br />
                <br />
                -- <a href="Asset_Details_Like.asp?SType=3">Lost Assets</a>
                <br />
                <br />
                -- <a href="Asset_Details_Like.asp?SType=4">Damaged Assets</a>
                <br />
                <br />
                -- <a href="Assets_Inv.asp?d=1">Assets's Inventory</a>
                <br />
                <br />
                -- <a href="Asset_Add.asp">Add New Asset</a>
                <br />
                <br />
                -- <a href="Emp_Add.asp">Add New Employee</a>
            </td>
            <td width='50'></td>
            <td>
                <h2>Last 5 History</h2>
                <table>
                    <tr>
                        <td class="TableTitle">Date</td>
                        <td class="TableTitle">Owner</td>
                        <td class="TableTitle">Note</td>
                    </tr>
                    <tr>
                        <td class="cell">23-06-2015 10:13:45</td>
                        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">11322</a>( Mohammed Al Jutail)</td>
                        <td class="cell">New asset assigned ( RC - 21232 ) Because he requested</td>
                    </tr>
                    <tr>
                        <td class="cell">23-06-2015 10:13:44</td>
                        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">11322</a>( Mohammed Al Jutail)</td>
                        <td class="cell">Asset Released ( RC - 12312 ) Because it is old</td>
                    </tr>
                    <tr>
                        <td class="cell">23-06-2015 09:08:33</td>
                        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">33246</a>( Abdulaziz Al Mutairi)</td>
                        <td class="cell">New asset assigned ( RS - 88774 ) Because he requested</td>
                    </tr>
                    <tr>
                        <td class="cell">23-06-2015 08:11:29</td>
                        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">65487</a>( Ali Al Saleh)</td>
                        <td class="cell">New asset assigned ( RC - 77454 ) Because he is new employee</td>
                    </tr>
                    <tr>
                        <td class="cell">23-06-2015 08:09:11</td>
                        <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">41231</a>( Salman Al rezq)</td>
                        <td class="cell">Asset Released ( RC - 78943 ) Because it is old</td>
                    </tr>
                </table>
                <table>
                    <tr>
                        <td>
                            <h2>Last 5 Employees</h2>
                            <table>
                                <tr>
                                    <td class="TableTitle">Date</td>
                                    <td class="TableTitle">Badge</td>
                                    <td class="TableTitle">Name</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 10:13:45</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">11322</a></td>
                                    <td class="cell">Mohammed Al Jutail</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 10:13:44</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">223398</a></td>
                                    <td class="cell">Ali Al Malek</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 09:08:33</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">33246</a></td>
                                    <td class="cell">Abdulaziz Al Mutairi</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 08:11:29</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">65487</a></td>
                                    <td class="cell">Ali Al Saleh</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 08:09:11</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">41231</a></td>
                                    <td class="cell">Salman Al rezq</td>
                                </tr>
                            </table>
                        </td>
                        <td>
                             &emsp;&emsp;
                        </td>
                        <td>
                            <h2>Last 5 Assets</h2>
                            <table>
                                <tr>
                                    <td class="TableTitle">Date</td>
                                    <td class="TableTitle">ISD Name</td>
                                    <td class="TableTitle">Type</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 10:13:45</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">RC - 31246</a></td>
                                    <td class="cell">PC</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 10:13:44</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">RL - 91221</a></td>
                                    <td class="cell">Laptop</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 09:08:33</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">RS - 65412</a></td>
                                    <td class="cell">Scanner</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 08:11:29</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">RI - 113</a></td>
                                    <td class="cell">iPad</td>
                                </tr>
                                <tr>
                                    <td class="cell">23-06-2015 08:09:11</td>
                                    <td class="cell"><a href="Emp_Details.asp?Emp_Badge=2">RP- 64321</a></td>
                                    <td class="cell">Printer</td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                </table>

            </td>
        </tr>
        <tr>
            <td>
                <table>
                    <tr>
                        <td>
                            <h2>Assets</h2>
                        </td>
                    </tr>
                    <tr>
                        <td class="TableTitle">Type</td>
                        <td class="TableTitle">Free</td>
                        <td class="TableTitle">Assigned</td>
                        <td class="TableTitle">Total</td>
                    </tr>
                    <tr>
                        <td class="cellT">PC</td>
                        <td class="cell">132</td>
                        <td class="cell">543</td>
                        <td class="cell">675</td>
                    </tr>
                    <tr>
                        <td class="cellT">LapTop</td>
                        <td class="cell">37</td>
                        <td class="cell">136</td>
                        <td class="cell">173</td>
                    </tr>
                    <tr>
                        <td class="cellT">Moniter</td>
                        <td class="cell">135</td>
                        <td class="cell">561</td>
                        <td class="cell">696</td>
                    </tr>
                    <tr>
                        <td class="cellT">Printer</td>
                        <td class="cell">73</td>
                        <td class="cell">192</td>
                        <td class="cell">265</td>
                    </tr>
                    <tr>
                        <td class="cellT">Scanner</td>
                        <td class="cell">19</td>
                        <td class="cell">38</td>
                        <td class="cell">57</td>
                    </tr>
                    <tr>
                        <td class="cellT">iPad</td>
                        <td class="cell">97</td>
                        <td class="cell">3</td>
                        <td class="cell">100</td>
                    </tr>
                </table>
            </td>
            <td width='50'></td>
            <td>
                <table>
                    <tr>
                        <td>
                            <h2>Percentage</h2>
                        </td>
                    </tr>
                    <tr>
                        <td class="TableTitle">Type</td>
                        <td class="TableTitle">Assigned &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;
    &emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp;&emsp; Free</td>
                    </tr>
                    <tr>
                        <td class="cellT">PC</td>
                        <td class="cell"><%Response.Write((FormatNumber((543/675) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((543/675) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(543/675)) * 100)))%>%</td>
                    </tr>
                    <tr>
                        <td class="cellT">LapTop</td>
                        <td class="cell"><%Response.Write((FormatNumber((136/173) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((136/173) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(136/173)) * 100)))%>%</td>
                    </tr>
                    <tr>
                        <td class="cellT">Moniter</td>
                        <td class="cell"><%Response.Write((FormatNumber((561/696) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((561/696) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(561/696)) * 100)))%>%</td>
                    </tr>
                    <tr>
                        <td class="cellT">Printer</td>
                        <td class="cell"><%Response.Write((FormatNumber((192/265) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((192/265) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(192/265)) * 100)))%>%</td>
                    </tr>
                    <tr>
                        <td class="cellT">Scanner</td>
                        <td class="cell"><%Response.Write((FormatNumber((38/57) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((38/57) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(38/57)) * 100)))%>%</td>
                    </tr>
                    <tr>
                        <td class="cellT">iPad</td>
                        <td class="cell">0<%Response.Write((FormatNumber((3/100) * 100)))%>%&emsp;
                <%ii = int (FormatNumber((3/100) * 100))
                i = 1
                %><b>[&nbsp;</b><%			
                Do While i < ii
                %><font color="red"><b>l</b></font><%
                    i = i + 1
                loop
                Do While i < 100
                %><font color="green"><b>l</b></font><%
                    i = i + 1
                loop
                %><b>&nbsp;]</b>&emsp;
                <%Response.Write((FormatNumber((1-(3/100)) * 100)))%>%</td>
                    </tr>
                </table>

            </td>
        </tr>
    </table>
</body>
</html>
<!-- #include file="footer.asp" -->
