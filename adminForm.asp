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
</head>
<body>
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
