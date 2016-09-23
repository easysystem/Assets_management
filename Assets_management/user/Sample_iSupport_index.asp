<html dir="ltr">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    <link href="modern.css" rel="stylesheet" type="text/css" />
    <link href="isupport1.css" rel="stylesheet" type="text/css" />
    <link href="MetroJs.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="carousel.js"></script>
    <script language="JavaScript" src="tile-slider.js"></script>
    <script language="JavaScript" src="MetroJs.js"></script>
    <link rel="stylesheet" type="text/css" href="style.css" />
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1256" />
    <script type="text/javascript">
        function showCustomer(str) {
            if (str == "") {
                document.getElementById("txtHint").innerHTML = "";
                return;
            }
            if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp = new XMLHttpRequest();
            }
            else {// code for IE6, IE5
                xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            xmlhttp.onreadystatechange = function () {
                if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                    document.getElementById("txtHint").innerHTML = xmlhttp.responseText;
                }
            }
            //alert("Relase_Assets.asp?q=" + str + "abc");
            xmlhttp.open("GET", "Relase_Assets.asp?q=" + str, true);
            xmlhttp.send();
        }
    </script>
    <script type="text/javascript">
        function NewAssets(ida) {
            if (ida == "") {
                document.getElementById("txtHint2").innerHTML = "";
                return;
            }
            if (window.XMLHttpRequest) {// code for IE7+, Firefox, Chrome, Opera, Safari
                xmlhttp = new XMLHttpRequest();
            }
            else {// code for IE6, IE5
                xmlhttp = new ActiveXObject("Microsoft.XMLHTTP");
            }
            xmlhttp.onreadystatechange = function () {
                if (xmlhttp.readyState == 4 && xmlhttp.status == 200) {
                    document.getElementById("txtHint2").innerHTML = xmlhttp.responseText;
                }
            }
            xmlhttp.open("GET", "New_Assets.asp?q=" + ida, true);
            xmlhttp.send();
        }
    </script>
    <script language="Javascript">

        function BYBadge() {
            if (document.search.Searchb.value == "") {
                alert("Sorry , You did not give me any digit from the Badge");
            }
            else {
                var numericExpression = /^[0-9]+$/;
                if (document.search.Searchb.value.match(numericExpression)) {
                    document.search.action = "Emp_Details.asp"
                    document.search.submit();             // Submit the page
                    return true;
                }
                else {
                    var answer = confirm("I think you want me to search by Name not Badge, aren't you?")
                    if (answer) {
                        document.search.action = "Emp_Details.asp?SType=2"
                        document.search.submit();             // Submit the page
                        return true;
                    }
                    else {

                        alert("Sorry , Badge it must be numbers only");
                        search.Searchb.focus();
                        return false;
                    }
                }
            }
        }

        function BYName() {
            document.search.action = "Emp_Details.asp?SType=2"
            document.search.submit();             // Submit the page
            return true;
        }

    </script>
</head>
<body class="bg-color-custom">
    <div class="bg-color-blue overflow: visible;">
        <form method="POST" action="Emp_Details.asp" name="search" onsubmit="return BYBadge()">
            <img style="margin-left: 10px; margin-right: 10px" height="25" src="images/Blogo.png" />
            <input name="Searchb" class="bg-color-wight ; fg-color-blue ; border-color-blue ; outline-color-blueDark" title="Searchb" value="" size="20" maxlength="50" placeholder="  Looking for..." />
            <input class="default" type="button" value="Search by BADGE" name="button1" onclick="return BYBadge();">
            <input class="default" type="button" value="Search by  NAME" name="button2" onclick="return BYName();">
            <button class="image-button bg-color-blue fg-color-white">
                <img src="images/home.png" /><a class="link" href="Emp_Details.asp?Emp_Badge=<% response.Write(Session("SUser_ID")) %>">Home</a></button>
            <button class="image-button bg-color-blue fg-color-white">
                <img src="images/logout.png" /><a class="link" href="logout.asp">Log Out</a></button>
            <font color="white" style="margin-left: 100px">Welcome Dear Mr. Mohammed Al - Jutail </font>

        </form>
    </div>
    
    <h1>
        Your information</h1>
    <table>
        <tr>
            <td>
                <table class="striped2">
                    <tr>
                        <td>Name:</td>
                        <td>Mohammed Al - Jutail</td>
                    </tr>
                    <tr>
                        <td>Tittle</td>
                        <td>Application Analist</td>
                    </tr>
                    <tr>
                        <td>Badge</td>
                        <td>112233</td>
                    </tr>
                    <tr>
                        <td>Email</td>
                        <td>es@es.sa</td>
                    </tr>
                    <tr>
                        <td>Department</td>
                        <td>Information tecnology</td>
                    </tr>
                    <tr>
                        <td>Area</td>
                        <td>IT </td>
                    </tr>
                    <tr>
                        <td>Office#</td>
                        <td>121 - A - 2</td>
                    </tr>
                    <tr>
                        <td>Ext.</td>
                        <td>40640</td>
                    </tr>
                    <tr>
                        <td>Pager</td>
                        <td>12121</td>
                    </tr>
                    <tr>
                        <td>Boss</td>
                        <td><a href="Emp_Details.asp?Emp_Badge=1">Dr. Mohammed Ali</a></td>
                    </tr>
                    <tr>
                        <td>Registered</td>
                        <td>10-02-2010</td>
                    </tr>
                    <tr>
                        <td>Note</td>
                        <td>New Employee</td>
                    </tr>
                </table>
            </td>
            <td width="350px">
                <div class="tile bg-color-red">
                    <div class="tile-content">
                        If you planning to arrange meeting and you need IT tech support or IT asset ,.. etc clik here
                    </div>
                    <div class="brand">
                        <div class="name">Support Meeting</div>
                    </div>
                </div>
                <div class="tile bg-color-green">
                    <div class="tile-content">
                        If you would like to request new IT asset such as PC, Laptop, Keyboard ,.. etc click here
                    </div>
                    <div class="brand">
                        <div class="name">New Asset</div>
                    </div>
                </div>
                <div class="tile double bg-color-blue" data-role="tile-slider" data-param-period="3000">
                    <div class="tile-content">
                        <h2>iSupport@***.sa</h2>
                        <h5>Leave a feedback</h5>
                        <p>
                            If you want to evaluate my services or suggest any thing , click here
                        </p>
                    </div>
                    <div class="tile-content">
                        <h2>iSupport@***.sa</h2>
                        <h5>Leave a feedback</h5>
                        <p>
                            If you want to evaluate my services or suggest any thing , click here
                        </p>
                    </div>
                    <div class="brand">
                        <div class="badge">12</div>
                        <img class="icon" src="images/home.png" />
                    </div>
                </div>
            </td>
        </tr>
    </table>
    <h2>Your Assets</h2>

    <table class="striped">
        <tbody>
            <tr>
                <th>Type</th>
                <th>ISD-Name</th>
                <th>Assigning Date</th>
                <th>Action</th>
            </tr>

            <tr>
                <td>PC</td>
                <td>RC - 10102</td>
                <td>20-02-2010</td>
                <td>
                    <label class="input-control switch">
                        <input type="checkbox">
                        <span class="helper">Confirm you have this asset to start supporting you</span>
                    </label>
                </td>
            </tr>
            <tr>
                <td>Laptop</td>
                <td>RL - 20301</td>
                <td>22-08-2013</td>
                <td>
                    <label class="input-control switch">
                        <input type="checkbox">
                        <span class="helper">Confirm you have this asset to start supporting you</span>
                    </label>
                </td>
            </tr>
            <tr>
                <td>Printer</td>
                <td>RR - 99070</td>
                <td>22-05-2011</td>
                <td>
                    <label class="input-control switch">
                        <input type="checkbox">
                        <span class="helper">Confirm you have this asset to start supporting you</span>
                    </label>
                </td>
            </tr>
        </tbody>
    </table>
    <div id="txtHint"></div>


    <h2>Your History</h2>

    <table class="striped">

        <tr>
            <th>Type</th>
            <th>ISD-Name</th>
            <th>Assigning Date</th>
            <th>Note</th>
        </tr>

        <tr>
            <td>PC</td>
            <td>RC - 66448</td>
            <td>20-02-2010</td>
            <td>Replaced because it is old one </td>
        </tr>
        <tr>
            <td>Printer</td>
            <td>RP - 15553</td>
            <td>09-08-2012</td>
            <td>Replaced because does not work any more</td>
        </tr>
        <tr>
            <td>Laptop</td>
            <td>RL - 66332</td>
            <td>11-03-2013</td>
            <td>Returned , no need any more</td>
        </tr>
        <tr>
            <td>Scanner</td>
            <td>RS - 11112</td>
            <td>21-12-2012</td>
            <td>ss</td>
        </tr>
    </table>

    <br />
    <br />
</body>
</html>
