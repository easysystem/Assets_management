
<html dir="ltr"  >
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1256"/>
    <link href="modern.css" rel="stylesheet" type="text/css" />
    <link href="isupport1.css" rel="stylesheet" type="text/css" />
    <link href="MetroJs.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="carousel.js"></script>
    <SCRIPT LANGUAGE="JavaScript" SRC="tile-slider.js"></SCRIPT>
    <SCRIPT LANGUAGE="JavaScript" SRC="MetroJs.js"></SCRIPT>
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

    function BYName()
    {
        document.search.action = "Emp_Details.asp?SType=2"
        document.search.submit();             // Submit the page
        return true;
    }

</script>
</head>
    <body class="bg-color-custom">	 
    <div class="bg-color-blue overflow: visible;">
        <form method="POST" action="Emp_Details.asp" name="search" onsubmit="return BYBadge()"  >
            <img style="margin-left : 10px ; margin-right:10px" height="25" src="images/Blogo.png"/>
            <input name="Searchb" class="bg-color-wight ; fg-color-blue ; border-color-blue ; outline-color-blueDark" title="Searchb" value="" size="20" maxlength="50" placeholder="  Looking for..." />
            <INPUT class="default" type="button" value="Search by BADGE" name=button1 onclick="return BYBadge();">
            <INPUT class="default" type="button" value="Search by  NAME" name=button2 onclick="return BYName();">
            <button class="image-button bg-color-blue fg-color-white"><img src="images/home.png"/><a class="link" href="Emp_Details.asp?Emp_Badge=<% response.Write(Session("SUser_ID")) %>">Home</a></button>
            <button class="image-button bg-color-blue fg-color-white"><img src="images/logout.png"/><a class="link" href="logout.asp">Log Out</a></button>
            <font color="white"style="margin-left : 100px">Welcome Dear <% response.Write(Session("SUser_name")) %> </font>
            
        </form>
    </div>
	</body>
</html>