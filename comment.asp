<%@ language="JavaScript" %>
<!DOCTYPE HTML>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>Comment</title>
    </head>
    <body>
        <%
            var Comment = Request.Form("commentin");
            var Time = new Date();
            var formattedTime = Time.getFullYear() + "-" + (Time.getMonth() + 1) + "-" + Time.getDate() + " " + Time.getHours() + ":" + Time.getMinutes() + ":" + Time.getSeconds();
            var MovieID = 1;
            var UserID = 1001;

            var conn = new ActiveXObject("ADODB.Connection");
            conn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='C:/Users/RossDAI/Documents/Movies.accdb'");
            var sqlString = "INSERT INTO Comment(UserID, MovieID, CommentTime, Comment) VALUES(" + UserID + ", " + MovieID + ", '" + formattedTime + "', '" + Comment + "')";
            conn.Execute(sqlString);
            conn.Close();
            Response.Redirect("Perfect_Days.html");
        %>
    </body>
</html>