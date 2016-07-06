<% 
    'START: variables
    Dim displayCode
    Dim displayText
    displayCode=Request.QueryString("displaySubmit")
    'END: variables
	
	'START: make cookie
	Response.Cookies("wayfinderAdministration").Expires = Date() + 7
	Response.Cookies("wayfinderAdministration")("displayCode") = displayCode
	'END: make cookie
	

    'START: prohibit page caching
    Response.Expires = -1000
    Response.Buffer = True
    'END: prohibit page caching

    'START: login redirect
    If Session("userLoggedIn") <> "true" And Request.ServerVariables("remote_addr") <> "192.168.220.84" And Request.ServerVariables("remote_addr") <> "192.168.220.85" Then
		Response.Redirect("default.asp")
	End If
    'END: login redirect

    'START: form processing
    If displayCode<>"" Then
        If displayCode = 0 Then
			Application("displayStatus") = displayCode
            displayText = "Normal message enabled."
			Response.Cookies("wayfinderAdministration")("codeNrml") = Now
        ElseIf displayCode = 1 Then
			Application("displayStatus") = displayCode
            displayText = "Evacuation message enabled."
			Response.Cookies("wayfinderAdministration")("codeEvac") = Now
		ElseIf displayCode = 2 Then
			Application("displayStatus") = displayCode
            displayText = "Emergency message enabled."
			Response.Cookies("wayfinderAdministration")("codeEmrg") = Now
		Else
			displayText = "Invalid entry."
        End If
    End If
    'END: form processing
%>
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Wayfinder Administration</title>
	<style type="text/css">
		body{font-family:sans-serif;font-size:10pt;}
		h1{font-size:18pt;margin:0px;padding:0px;}
		.topmenu{font-size:9pt;padding:4px 0px 20px 0px;}
	</style>

</head>
<body>
	<h1>Wayfinder Administration</h1>
	<div class="topmenu"><a href="adminForm.asp">administration</a> | <a href="display.asp">directory</a> | <a href="readme.asp">readme</a></div>
	<form id="adminForm" method="get" action="adminForm.asp">
		<select name="displaySubmit">
			<option value="0">Normal Status</option>
			<option value="1">Evacuation Status</option>			
			<option value="2">Emergency Status</option>
		</select>
		<input type="submit" id="submit" />
		<br />
        <%= displayText %>
	</form>
</body>
</html>
