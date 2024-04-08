<%@ language="JavaScript" %>
<!DOCTYPE HTML>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Sign Up</title>
    </head>
    <body>
        <%
            var UserID = Request.Form("UserID");
            var UserName = Request.Form("UserName");
            var Mail = Request.Form("Mail");
            var Gender = Request.Form("Gender");

            var conn = new ActiveXObject("ADODB.Connection");
            conn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='C:/Users/RossDAI/Documents/Movies.accdb'");
            var sqlString = "INSERT INTO [User]([UserID], [UserName], [Gender], [Mail]) VALUES('" + UserID + "', '" + UserName + "', '" + Gender + "', '" + Mail + "')";
            conn.Execute(sqlString);
            conn.Close();
            Response.Redirect("Homepage.html");
        %>
    </body>
</html>