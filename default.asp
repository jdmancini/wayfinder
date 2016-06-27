<%
    'START: variables
    Dim username
    Dim password
    username = ""  ' change this to set the username
    password = "" ' change this to set the password
    'END: variables


    'START: prohibit page caching
    Response.Expires = -1000
    Response.Buffer = True
	Session.Timeout = 3 'session times out in 3 minutes with no activity
    'END: prohibit page caching
    
    'START: login form and processing
    If Request.QueryString("fLogin") = "true" Then 
        CheckLogin 
    Else
        ShowLogin 
    End If 
    Sub ShowLogin
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Wayfinder Administration</title>
</head>
<body>
    <form name="loginForm" method="get" action="default.asp">
        <input type="text" name="fUsername">
        <input type="password" name="fPassword">
        <input type="hidden" name="fLogin" value="true">
        <input type="submit" value="submit">
    </form>
</body>
</html>
<%
    End Sub 
    Sub CheckLogin
    If LCase(Request.QueryString("fUsername")) = username And LCase(Request.QueryString("fPassword")) = password Then
        Session("userLoggedIn") = "true"
        Response.Redirect "adminForm.asp"
    Else
        ShowLogin
		Response.Write("Login Failed.")
    End If
    End Sub 
    'END: login form and processing
%>