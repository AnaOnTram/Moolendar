<%@language="JavaScript" %>
<!DOCTYPE html>
<html lang="en">
    <head>
        <meta charset="utf-8" />
        <title>MovieRecords</title>
    </head>
    <body>
        <%
            var conn = new ActiveXObject("ADODB.Connection");
            var rs = new ActiveXObject("ADODB.Recordset");
            try {
                conn.Open("Provider=Microsoft.ACE.OLEDB.12.0; Data Source='C:/Users/RossDAI/Documents/Movies.accdb'");
                rs.Open("SELECT * FROM Movie", conn);
                while (!rs.EOF) {
                    Response.Write(rs.Fields("Movie Name").Value + ":" + rs.Fields("Director").Value + "<br>");
                    rs.MoveNext();
                }
            } catch (e) {
                Response.Write("Error: " + e.description);
            } finally {
                if (!rs.EOF) {
                    rs.Close();
                }
                if (conn.State === 1) {
                    conn.Close();
                }
            }
        %>
    </body>
</html>